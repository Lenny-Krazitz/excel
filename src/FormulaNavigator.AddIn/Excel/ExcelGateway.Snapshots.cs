using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using FormulaNavigator.Core;

namespace FormulaNavigator.AddIn.Excel
{
    public sealed partial class ExcelGateway
    {
        private async void BuildSnapshotAsync(WorkbookIndex state, List<InspectionContext> formulas,
            CancellationTokenSource generation)
        {
            CancellationToken cancellation = generation.Token;
            try
            {
                // The input contains strings and locations only, never Excel RCWs.
                DependencySnapshot snapshot = await Task.Run(() => ParseSnapshot(formulas, cancellation), cancellation)
                    .ConfigureAwait(false);
                await OnExcelThread(() =>
                {
                    if (!state.Closed && ReferenceEquals(state.PlanCancellation, generation))
                    {
                        string schema = ReadSchema(state);
                        if (state.Closed || !ReferenceEquals(state.PlanCancellation, generation)) return false;
                        if (state.Schema != schema) InvalidatePlan(state);
                        else
                        {
                            state.Snapshot = snapshot;
                            if (snapshot.DynamicFormulas.Count == 0) state.Dynamic = EmptyResolvedDependencies;
                            state.Parsing = false;
                            state.Progress = "Формулы проиндексированы: " + snapshot.Formulas.Count + ".";
                        }
                    }
                    return true;
                }).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { }
            catch (Exception error)
            {
                await RecordBuildError(state, generation, false, error).ConfigureAwait(false);
            }
        }

        private async Task RecordBuildError(WorkbookIndex state, CancellationTokenSource generation,
            bool dynamic, Exception error)
        {
            try
            {
                await OnExcelThread(() =>
                {
                    if (!state.Closed && ReferenceEquals(dynamic ? state.DynamicCancellation : state.PlanCancellation, generation))
                    {
                        state.Error = error;
                        state.Parsing = false;
                        state.DynamicBuilding = false;
                    }
                    return true;
                }).ConfigureAwait(false);
            }
            catch (Exception) { /* The dispatcher can disappear while the add-in shuts down. */ }
        }

        private static DependencySnapshot ParseSnapshot(List<InspectionContext> formulas, CancellationToken cancellation)
        {
            var references = new List<IndexedDependency>();
            var dynamicFormulas = new List<DynamicFormula>();
            var warnings = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < formulas.Count; i++)
            {
                cancellation.ThrowIfCancellationRequested();
                InspectionContext context = formulas[i];
                FormulaParseResult parsed = FormulaParser.Parse(context.Formula);
                if (!parsed.Success || parsed.Root == null)
                {
                    AddIndexWarning(warnings, "Некоторые формулы не удалось разобрать; список зависимых может быть неполным.");
                    continue;
                }
                FormulaDependencyPlan plan = FormulaDependencyPlan.Create(parsed.Root,
                    context.Location.Workbook, context.Location.Sheet);
                foreach (ReferenceArea area in plan.References) references.Add(new IndexedDependency(area, i));
                foreach (string warning in plan.Warnings) AddIndexWarning(warnings, warning);
                if (plan.DynamicNodes.Count != 0) dynamicFormulas.Add(new DynamicFormula(i, context, plan.DynamicNodes));
            }
            cancellation.ThrowIfCancellationRequested();
            var index = new DependencyRangeIndex(references);
            cancellation.ThrowIfCancellationRequested();
            return new DependencySnapshot(formulas.AsReadOnly(), index, dynamicFormulas.AsReadOnly(), warnings.ToArray());
        }

        private static void AddIndexWarning(HashSet<string> warnings, string warning)
        {
            if (String.IsNullOrWhiteSpace(warning)) return;
            if (warnings.Count < 64) warnings.Add(warning);
            else warnings.Add("Другие неподдерживаемые ссылки также пропущены.");
        }

        private sealed class DependencySnapshot
        {
            public DependencySnapshot(IReadOnlyList<InspectionContext> formulas, DependencyRangeIndex index,
                IReadOnlyList<DynamicFormula> dynamicFormulas, IReadOnlyList<string> warnings)
            {
                Formulas = formulas;
                Index = index;
                DynamicFormulas = dynamicFormulas;
                Warnings = warnings;
            }
            public readonly IReadOnlyList<InspectionContext> Formulas;
            public readonly DependencyRangeIndex Index;
            public readonly IReadOnlyList<DynamicFormula> DynamicFormulas;
            public readonly IReadOnlyList<string> Warnings;
        }

        private sealed class DynamicFormula
        {
            public DynamicFormula(int id, InspectionContext context, IReadOnlyList<FormulaNode> nodes)
            {
                Id = id;
                Context = context;
                Nodes = nodes;
            }
            public readonly int Id;
            public readonly InspectionContext Context;
            public readonly IReadOnlyList<FormulaNode> Nodes;
        }

        private sealed class FormulaSnapshotReader
        {
            private readonly dynamic _workbook;
            private readonly string _workbookName;
            private readonly int _sheetCount;
            private int _sheetIndex = 1;
            private dynamic _formulaCells;
            private int _areaIndex;
            private int _areaCount;
            private dynamic _area;
            private int _areaRow;
            private int _areaColumn;
            private IEnumerator<FormulaReadBlock> _blocks;

            public FormulaSnapshotReader(object workbook, string name)
            {
                _workbook = workbook;
                _workbookName = name;
                _sheetCount = Convert.ToInt32(_workbook.Worksheets.Count, CultureInfo.InvariantCulture);
            }

            public readonly List<InspectionContext> Formulas = new List<InspectionContext>();
            public string SheetName { get; private set; }

            // Each call prepares one sheet/area or copies at most 4096 cells. Parsing happens later.
            public bool ReadNext()
            {
                if (_blocks != null)
                {
                    if (_blocks.MoveNext())
                    {
                        FormulaReadBlock next = _blocks.Current;
                        dynamic range = _area.Cells[next.Row, next.Column].Resize[next.Height, next.Width];
                        object values = ReadFormulas(range);
                        CopyFormulas(values, _areaRow + next.Row - 1, _areaColumn + next.Column - 1,
                            next.Height, next.Width);
                        return true;
                    }
                    _blocks.Dispose();
                    _blocks = null;
                    _area = null;
                }
                if (_formulaCells != null && _areaIndex <= _areaCount)
                {
                    _area = _formulaCells.Areas[_areaIndex++];
                    _areaRow = Convert.ToInt32(_area.Row, CultureInfo.InvariantCulture);
                    _areaColumn = Convert.ToInt32(_area.Column, CultureInfo.InvariantCulture);
                    int rows = Convert.ToInt32(_area.Rows.Count, CultureInfo.InvariantCulture);
                    int columns = Convert.ToInt32(_area.Columns.Count, CultureInfo.InvariantCulture);
                    _blocks = FormulaReadBlocks.Enumerate(rows, columns, MaxCellsPerRead).GetEnumerator();
                    return true;
                }
                if (_sheetIndex > _sheetCount) return false;
                dynamic sheet = _workbook.Worksheets[_sheetIndex++];
                SheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);
                _formulaCells = null;
                _areaIndex = 1;
                _areaCount = 0;
                try
                {
                    _formulaCells = sheet.Cells.SpecialCells(XlCellTypeFormulas);
                    _areaCount = Convert.ToInt32(_formulaCells.Areas.Count, CultureInfo.InvariantCulture);
                }
                catch (COMException error) when (unchecked((uint)error.ErrorCode) == 0x800A03EC)
                {
                    // SpecialCells reports this Excel error for worksheets with no formulas.
                }
                return true;
            }

            private void CopyFormulas(object values, int row, int column, int height, int width)
            {
                Array matrix = values as Array;
                if (matrix == null)
                {
                    AddFormula(values as string, row, column);
                    return;
                }
                int lowerRow = matrix.GetLowerBound(0);
                int lowerColumn = matrix.GetLowerBound(1);
                for (int r = 0; r < height; r++)
                    for (int c = 0; c < width; c++)
                        AddFormula(matrix.GetValue(lowerRow + r, lowerColumn + c) as string, row + r, column + c);
            }

            private void AddFormula(string formula, int row, int column)
            {
                if (!IsFormula(formula)) return;
                var location = new ExcelLocation(_workbookName, SheetName,
                    ColumnName(column) + row.ToString(CultureInfo.InvariantCulture));
                Formulas.Add(new InspectionContext(location, formula, String.Empty, true));
            }
        }
    }
}
