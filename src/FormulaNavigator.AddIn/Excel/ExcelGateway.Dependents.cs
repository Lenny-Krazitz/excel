using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using FormulaNavigator.Core;

namespace FormulaNavigator.AddIn.Excel
{
    public sealed partial class ExcelGateway
    {
        private static readonly ResolvedDependencies EmptyResolvedDependencies =
            new ResolvedDependencies(new DependencyRangeIndex(new IndexedDependency[0]), new string[0]);

        public async Task<DependencyResult> FindDependentsAsync(ExcelLocation location, CancellationToken cancellation,
            IProgress<string> progress, bool forceRefresh = false)
        {
            if (location == null) throw new ArgumentNullException(nameof(location));
            ReferenceArea source;
            if (!ReferenceParser.TryParse(location.Address, location.Workbook, location.Sheet, out source))
                throw new InvalidOperationException("Выбранный адрес не является поддерживаемым A1-диапазоном.");
            var elapsed = Stopwatch.StartNew();
            bool reused = false;
            string refreshWarning = null;
            WorkbookIndex state = await OnExcelThread(() =>
            {
                cancellation.ThrowIfCancellationRequested();
                StartIndexing();
                ReconcileWorkbooks();
                object workbook = FindOpenWorkbook(location.Workbook);
                WorkbookIndex session = _indexes[ComIdentity(workbook)];
                string schema = ReadSchema(session);
                bool canReuse = _eventsConnected && Convert.ToBoolean(_application.EnableEvents, CultureInfo.InvariantCulture);
                if (forceRefresh || !canReuse || session.SchemaWarning != null || session.Error != null || session.Schema != schema)
                    InvalidatePlan(session);
                if (!canReuse)
                    refreshWarning = "События Excel недоступны: книга перечитана полностью. После правок с отключёнными событиями используйте «Обновить».";
                session.Schema = schema;
                reused = session.Snapshot != null && session.Dynamic != null;
                session.Demand++;
                _priorityWorkbook = ComIdentity(workbook);
                return session;
            }).ConfigureAwait(false);

            try
            {
                string lastProgress = null;
                while (true)
                {
                    cancellation.ThrowIfCancellationRequested();
                    QuerySnapshot query = await OnExcelThread(() =>
                    {
                        cancellation.ThrowIfCancellationRequested();
                        if (state.Closed) throw new InvalidOperationException("Книга закрыта. Откройте окно поиска заново.");
                        if (state.Error != null) throw new InvalidOperationException("Не удалось построить индекс: " + ComMessage(state.Error), state.Error);
                        if (!IsCalculationStable())
                        {
                            state.Progress = "Ожидаю завершения пересчёта Excel…";
                            if (lastProgress != state.Progress) progress?.Report(state.Progress);
                            lastProgress = state.Progress;
                            return null;
                        }
                        if (state.Snapshot != null && state.Dynamic != null)
                        {
                            string calculationWarning = state.Snapshot.DynamicFormulas.Count != 0 &&
                                Convert.ToInt32(_application.CalculationState, CultureInfo.InvariantCulture) == 2
                                ? "Ручной пересчёт Excel: вычисляемые ссылки используют текущие рассчитанные значения. Для их обновления нажмите F9 в Excel."
                                : null;
                            return new QuerySnapshot(state.Snapshot, state.Dynamic, state.SchemaWarning, calculationWarning);
                        }
                        if (state.Progress != lastProgress)
                        {
                            lastProgress = state.Progress;
                            progress?.Report(lastProgress ?? "Подготавливаю индекс…");
                        }
                        return null;
                    }).ConfigureAwait(false);
                    if (query == null)
                    {
                        await Task.Delay(60, cancellation).ConfigureAwait(false);
                        continue;
                    }

                    progress?.Report("Ищу зависимые в готовом индексе…");
                    DependencyResult result = await Task.Run(() =>
                    {
                        var ids = new SortedSet<int>(query.Plan.Index.Find(source));
                        ids.UnionWith(query.Dynamic.Index.Find(source));
                        var cells = new List<DependentCell>();
                        foreach (int id in ids)
                        {
                            cancellation.ThrowIfCancellationRequested();
                            cells.Add(new DependentCell(query.Plan.Formulas[id], String.Empty));
                        }
                        var warnings = new HashSet<string>(query.Plan.Warnings, StringComparer.OrdinalIgnoreCase);
                        foreach (string warning in query.Dynamic.Warnings) AddIndexWarning(warnings, warning);
                        AddIndexWarning(warnings, refreshWarning);
                        AddIndexWarning(warnings, query.SchemaWarning);
                        AddIndexWarning(warnings, query.CalculationWarning);
                        return new DependencyResult(cells.AsReadOnly(), warnings.ToArray(), query.Plan.Formulas.Count,
                            reused, elapsed.ElapsedMilliseconds);
                    }, cancellation).ConfigureAwait(false);

                    bool current = await OnExcelThread(() =>
                    {
                        cancellation.ThrowIfCancellationRequested();
                        if (state.Closed) throw new InvalidOperationException("Книга закрыта.");
                        if (!ReferenceEquals(state.Snapshot, query.Plan) || !ReferenceEquals(state.Dynamic, query.Dynamic)) return false;
                        string schema = ReadSchema(state);
                        if (!ReferenceEquals(state.Snapshot, query.Plan) || !ReferenceEquals(state.Dynamic, query.Dynamic)) return false;
                        if (state.Schema != schema)
                        {
                            InvalidatePlan(state);
                            return false;
                        }
                        bool stable = IsCalculationStable();
                        return stable && ReferenceEquals(state.Snapshot, query.Plan) && ReferenceEquals(state.Dynamic, query.Dynamic);
                    }).ConfigureAwait(false);
                    if (current) return new DependencyResult(result.Cells, result.Warnings, result.FormulaCount,
                        result.IndexReused, elapsed.ElapsedMilliseconds);
                    reused = false;
                }
            }
            finally
            {
                try { await OnExcelThread(() => --state.Demand).ConfigureAwait(false); }
                catch (ObjectDisposedException) { }
                catch (OperationCanceledException) { }
            }
        }

        private void PumpDynamicIndex(WorkbookIndex state)
        {
            if (state.DynamicReader == null) state.DynamicReader = new DynamicSnapshotReader(state.Snapshot, state.Workbook);
            DynamicSnapshotReader reader = state.DynamicReader;
            CancellationTokenSource generation = state.DynamicCancellation;
            var elapsed = Stopwatch.StartNew();
            for (int count = 0; count < 16 && elapsed.ElapsedMilliseconds < 12; count++)
            {
                bool more = reader.ReadNext(this);
                // Evaluation can pump Excel messages and deliver a calculation/change event.
                if (!ReferenceEquals(state.DynamicCancellation, generation)) return;
                if (!more)
                {
                    state.DynamicReader = null;
                    state.DynamicBuilding = true;
                    BuildDynamicIndexAsync(state, reader, generation);
                    return;
                }
            }
            state.Progress = "Уточняю вычисляемые ссылки: " + reader.FormulaPosition + "/"
                + state.Snapshot.DynamicFormulas.Count + " формул…";
        }

        private async void BuildDynamicIndexAsync(WorkbookIndex state, DynamicSnapshotReader reader,
            CancellationTokenSource generation)
        {
            CancellationToken cancellation = generation.Token;
            // Do not capture the COM-bearing reader in the worker delegate.
            List<IndexedDependency> references = reader.References;
            string[] warnings = reader.Warnings.ToArray();
            try
            {
                ResolvedDependencies result = await Task.Run(() =>
                {
                    cancellation.ThrowIfCancellationRequested();
                    var index = new DependencyRangeIndex(references);
                    cancellation.ThrowIfCancellationRequested();
                    return new ResolvedDependencies(index, warnings);
                }, cancellation).ConfigureAwait(false);
                await OnExcelThread(() =>
                {
                    if (!state.Closed && ReferenceEquals(state.DynamicCancellation, generation))
                    {
                        bool stable = IsCalculationStable();
                        string schema = ReadSchema(state);
                        if (state.Closed || !ReferenceEquals(state.DynamicCancellation, generation)) return false;
                        if (!stable) InvalidateDynamic(state);
                        else if (state.Schema != schema) InvalidatePlan(state);
                        else
                        {
                            state.Dynamic = result;
                            state.DynamicBuilding = false;
                        }
                    }
                    return true;
                }).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { }
            catch (Exception error) { await RecordBuildError(state, generation, true, error).ConfigureAwait(false); }
        }

        private bool TryResolveDynamicAreas(InspectionContext context, object worksheet, FormulaNode node,
            out List<ReferenceArea> areas, out string note)
        {
            int budget = 16384;
            return TryResolveDynamicAreas(context, worksheet, node, ref budget, out areas, out note);
        }

        private bool TryResolveDynamicAreas(InspectionContext context, object worksheet, FormulaNode node,
            ref int budget, out List<ReferenceArea> areas, out string note)
        {
            areas = new List<ReferenceArea>();
            note = String.Empty;
            if (--budget < 0)
            {
                note = "Слишком сложное пересечение или объединение диапазонов пропущено.";
                return false;
            }
            if (node.Kind == FormulaNodeKind.Group && node.Children.Count == 1)
                return TryResolveDynamicAreas(context, worksheet, node.Children[0], ref budget, out areas, out note);
            if (node.Kind == FormulaNodeKind.Binary && node.Children.Count == 2 &&
                (node.Operator == " " || node.Operator == "," || node.Operator == ";"))
            {
                List<ReferenceArea> left, right;
                if (!TryResolveDynamicAreas(context, worksheet, node.Children[0], ref budget, out left, out note)
                    || !TryResolveDynamicAreas(context, worksheet, node.Children[1], ref budget, out right, out note)) return false;
                if (node.Operator != " ")
                {
                    areas.AddRange(left);
                    areas.AddRange(right);
                    return true;
                }
                foreach (ReferenceArea first in left)
                    foreach (ReferenceArea second in right)
                    {
                        if (--budget < 0)
                        {
                            note = "Слишком сложное пересечение диапазонов пропущено.";
                            return false;
                        }
                        if (first.Intersects(second))
                            areas.Add(new ReferenceArea(first.Workbook, first.Sheet,
                                Math.Max(first.StartRow, second.StartRow), Math.Min(first.EndRow, second.EndRow),
                                Math.Max(first.StartColumn, second.StartColumn), Math.Min(first.EndColumn, second.EndColumn)));
                    }
                return true;
            }
            Resolution resolution;
            if (!TryResolveNodeAsRange(context, worksheet, node, out resolution, out note))
            {
                if (String.IsNullOrEmpty(note)) note = "Не удалось определить адрес выражения: " + node.Text;
                return false;
            }
            dynamic parts = resolution.Range.Areas;
            int count = Convert.ToInt32(parts.Count, CultureInfo.InvariantCulture);
            budget -= count;
            if (budget < 0)
            {
                note = "Диапазон с чрезмерным количеством областей пропущен.";
                return false;
            }
            if (count == 1) areas.Add(resolution.Area);
            else
            {
                // Names may refer to a non-contiguous range. Its bounding box would create false matches.
                for (int i = 1; i <= count; i++)
                {
                    dynamic part = parts[i];
                    ReferenceArea area;
                    if (!ReferenceParser.TryParse(GetAddress(part), resolution.Area.Workbook, resolution.Area.Sheet, out area))
                    {
                        note = "Не удалось прочитать составной диапазон: " + node.Text;
                        return false;
                    }
                    areas.Add(area);
                }
            }
            return true;
        }

        private sealed class DynamicSnapshotReader
        {
            private readonly DependencySnapshot _snapshot;
            private readonly dynamic _workbook;
            private int _formula;
            private int _node;
            private string _sheetName;
            private object _worksheet;

            public DynamicSnapshotReader(DependencySnapshot snapshot, object workbook)
            {
                _snapshot = snapshot;
                _workbook = workbook;
            }
            public readonly List<IndexedDependency> References = new List<IndexedDependency>();
            public readonly HashSet<string> Warnings = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public int FormulaPosition { get { return _formula; } }

            public bool ReadNext(ExcelGateway gateway)
            {
                if (_formula >= _snapshot.DynamicFormulas.Count) return false;
                DynamicFormula formula = _snapshot.DynamicFormulas[_formula];
                if (_sheetName != formula.Context.Location.Sheet)
                {
                    _sheetName = formula.Context.Location.Sheet;
                    _worksheet = _workbook.Worksheets[_sheetName];
                }
                List<ReferenceArea> areas;
                string note;
                if (gateway.TryResolveDynamicAreas(formula.Context, _worksheet, formula.Nodes[_node], out areas, out note))
                    foreach (ReferenceArea area in areas) References.Add(new IndexedDependency(area, formula.Id));
                else AddIndexWarning(Warnings, note);
                if (++_node == formula.Nodes.Count) { _node = 0; _formula++; }
                return true;
            }
        }

        private sealed class ResolvedDependencies
        {
            public ResolvedDependencies(DependencyRangeIndex index, IReadOnlyList<string> warnings)
            {
                Index = index;
                Warnings = warnings;
            }
            public readonly DependencyRangeIndex Index;
            public readonly IReadOnlyList<string> Warnings;
        }

        private sealed class QuerySnapshot
        {
            public QuerySnapshot(DependencySnapshot plan, ResolvedDependencies dynamic, string schemaWarning, string calculationWarning)
            {
                Plan = plan;
                Dynamic = dynamic;
                SchemaWarning = schemaWarning;
                CalculationWarning = calculationWarning;
            }
            public readonly DependencySnapshot Plan;
            public readonly ResolvedDependencies Dynamic;
            public readonly string SchemaWarning;
            public readonly string CalculationWarning;
        }
    }
}
