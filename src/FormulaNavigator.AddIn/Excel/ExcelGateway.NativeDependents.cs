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
        // Outcome of the native dependents search: either a complete result built
        // from the Excel tracer arrows, or a warning explaining why the reverse
        // index must be used instead.
        private sealed class NativeSearchOutcome
        {
            public NativeSearchOutcome(DependencyResult result, string warning)
            {
                Result = result;
                Warning = warning;
            }

            public readonly DependencyResult Result;
            public readonly string Warning;
        }

        private sealed class NativeCellData
        {
            public readonly string Formula;
            public readonly object Value;

            public NativeCellData(string formula, object value)
            {
                Formula = formula;
                Value = value;
            }
        }

        // Attempts the built-in Excel tracer search (ShowDependents + NavigateArrow)
        // for the direct dependents of the cell. The COM work runs on Excel's STA
        // thread one arrow at a time, so the user can cancel while it is running.
        // Any failure, an unconfirmable result or a too large result returns null
        // and a warning; the caller then searches through the reverse index.
        private async Task<NativeSearchOutcome> TryFindNativeDependentsAsync(ExcelLocation location,
            CancellationToken cancellation, IProgress<string> progress)
        {
            var clock = Stopwatch.StartNew();
            var host = new ExcelNativeTraceHost(this, location);
            Interlocked.Increment(ref _nativeTraceActive);
            try
            {
                progress?.Report("Выполняю штатную трассировку Excel…");
                NativeTraceAttempt attempt = await NativeDependentTrace.RunAsync(host, location.Workbook, cancellation)
                    .ConfigureAwait(false);
                NativeTraceDecision decision = NativeDependentTrace.Decide(attempt);
                if (!decision.UseNative) return new NativeSearchOutcome(null, decision.Warning);

                IReadOnlyList<NativeTraceArea> merged = NativeDependentTrace.MergeAreas(attempt.Areas);
                IReadOnlyList<NativeTraceCell> cells;
                if (!NativeDependentTrace.TryExpandCells(merged, NativeDependentTrace.MaxCells, out cells))
                    return new NativeSearchOutcome(null,
                        "Штатная трассировка вернула слишком большой диапазон; выполнен поиск по индексу.");

                List<DependentCell> resultCells = await ReadNativeCellsAsync(merged, cells, cancellation)
                    .ConfigureAwait(false);
                clock.Stop();
                return new NativeSearchOutcome(new DependencyResult(resultCells.AsReadOnly(), new string[0],
                    resultCells.Count, false, clock.ElapsedMilliseconds, true), null);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception error)
            {
                return new NativeSearchOutcome(null,
                    "Штатная трассировка прервана ошибкой: " + ComMessage(error) + ". Выполнен поиск по индексу.");
            }
            finally
            {
                Interlocked.Decrement(ref _nativeTraceActive);
            }
        }

        // Reads formulas and values only for the found dependent areas and maps
        // them onto the ordered cell list produced by the trace.
        private async Task<List<DependentCell>> ReadNativeCellsAsync(IReadOnlyList<NativeTraceArea> areas,
            IReadOnlyList<NativeTraceCell> cells, CancellationToken cancellation)
        {
            var data = new Dictionary<string, NativeCellData>(StringComparer.OrdinalIgnoreCase);
            foreach (NativeTraceArea area in areas)
            {
                cancellation.ThrowIfCancellationRequested();
                Dictionary<string, NativeCellData> read = await OnExcelThread(() => ReadAreaData(area))
                    .ConfigureAwait(false);
                foreach (KeyValuePair<string, NativeCellData> pair in read) data[pair.Key] = pair.Value;
            }
            var result = new List<DependentCell>();
            foreach (NativeTraceCell cell in NativeDependentTrace.OrderCells(cells))
            {
                NativeCellData found;
                if (!data.TryGetValue(cell.Workbook + "|" + cell.Sheet + "|" + cell.Address, out found)) continue;
                var location = new ExcelLocation(cell.Workbook, cell.Sheet, cell.Address);
                string formula = found.Formula ?? String.Empty;
                var context = new InspectionContext(location, formula, DisplayValue(found.Value), IsFormula(formula));
                result.Add(new DependentCell(context, String.Empty));
            }
            return result;
        }

        private Dictionary<string, NativeCellData> ReadAreaData(NativeTraceArea area)
        {
            EnsureExcelThread();
            dynamic workbook = FindOpenWorkbook(area.Workbook);
            dynamic sheet = workbook.Worksheets[area.Sheet];
            dynamic range = sheet.Range[NativeDependentTrace.NormalizeAddress(area.Address)];
            var data = new Dictionary<string, NativeCellData>(StringComparer.OrdinalIgnoreCase);
            object formulas = ReadFormulas(range);
            object values;
            try { values = range.Value2; }
            catch (Exception) { values = null; }
            int firstRow = Convert.ToInt32(range.Row, CultureInfo.InvariantCulture);
            int firstColumn = Convert.ToInt32(range.Column, CultureInfo.InvariantCulture);
            int rowCount = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
            int columnCount = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
            for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
            {
                for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
                {
                    string address = ColumnName(firstColumn + columnOffset)
                        + (firstRow + rowOffset).ToString(CultureInfo.InvariantCulture);
                    string key = area.Workbook + "|" + area.Sheet + "|" + address;
                    data[key] = new NativeCellData(
                        Convert.ToString(ArrayCell(formulas, rowOffset, columnOffset), CultureInfo.InvariantCulture),
                        ArrayCell(values, rowOffset, columnOffset));
                }
            }
            return data;
        }

        // Excel returns either a scalar (single-cell range) or a 1-based 2D array.
        private static object ArrayCell(object values, int rowOffset, int columnOffset)
        {
            Array array = values as Array;
            if (array == null) return values;
            if (array.Rank != 2) return null;
            int row = Math.Min(rowOffset + array.GetLowerBound(0), array.GetUpperBound(0));
            int column = Math.Min(columnOffset + array.GetLowerBound(1), array.GetUpperBound(1));
            return array.GetValue(row, column);
        }

        // All tracer work runs on Excel's STA thread through the dispatcher. Every
        // step reports success or a failure reason; the trace never touches arrows
        // before the user-arrow check confirms that the cell has none of its own.
        private sealed class ExcelNativeTraceHost : INativeTraceHost
        {
            private readonly ExcelGateway _gateway;
            private readonly ExcelLocation _source;
            private readonly StateRestoreStack _restore = new StateRestoreStack();
            private readonly List<string> _cleanupFailures = new List<string>();
            private dynamic _workbook;
            private dynamic _sheet;
            private dynamic _sourceRange;
            private dynamic _originalWorkbook;
            private dynamic _originalSheet;
            private dynamic _originalSelection;
            private dynamic _originalActiveCell;
            private bool _screenUpdating;
            private bool _enableEvents;
            private bool _displayAlerts;
            private string _lastFailure;

            public ExcelNativeTraceHost(ExcelGateway gateway, ExcelLocation source)
            {
                _gateway = gateway;
                _source = source;
            }

            public IReadOnlyList<string> CleanupFailures { get { return _cleanupFailures.AsReadOnly(); } }
            public string LastFailure { get { return _lastFailure; } }

            public Task<bool> TrySaveStateAsync()
            {
                return _gateway.OnExcelThread(() => SaveState());
            }

            public Task<UserArrowCheck> CheckUserArrowsAsync()
            {
                return _gateway.OnExcelThread(() => CheckUserArrows());
            }

            public Task<bool> TryCreateArrowsAsync()
            {
                return _gateway.OnExcelThread(() => CreateArrows());
            }

            public Task<NativeArrowStep> NavigateAsync(int arrowNumber, int? linkNumber)
            {
                return _gateway.OnExcelThread(() => NavigateStep(arrowNumber, linkNumber));
            }

            public Task RemoveCreatedArrowsAsync()
            {
                return _gateway.OnExcelThread(() => RemoveCreatedArrows());
            }

            public Task RestoreStateAsync()
            {
                return _gateway.OnExcelThread(() => RestoreState());
            }

            private bool Fail(string message)
            {
                _lastFailure = message;
                return false;
            }

            // Reads the complete state snapshot before anything is modified. A
            // failure here aborts the native path without changing Excel at all.
            private bool SaveState()
            {
                dynamic application = _gateway._application;
                try { _originalWorkbook = application.ActiveWorkbook; }
                catch (Exception error) { return Fail("не удалось сохранить активную книгу: " + ComMessage(error)); }
                try { _originalSheet = application.ActiveSheet; }
                catch (Exception error) { return Fail("не удалось сохранить активный лист: " + ComMessage(error)); }
                try { _originalSelection = application.Selection; }
                catch (Exception error) { return Fail("не удалось сохранить выделение: " + ComMessage(error)); }
                try { _originalActiveCell = application.ActiveCell; }
                catch (Exception error) { return Fail("не удалось сохранить активную ячейку: " + ComMessage(error)); }
                try { _screenUpdating = Convert.ToBoolean(application.ScreenUpdating, CultureInfo.InvariantCulture); }
                catch (Exception error) { return Fail("не удалось сохранить настройку обновления экрана: " + ComMessage(error)); }
                try { _enableEvents = Convert.ToBoolean(application.EnableEvents, CultureInfo.InvariantCulture); }
                catch (Exception error) { return Fail("не удалось сохранить настройку событий: " + ComMessage(error)); }
                try { _displayAlerts = Convert.ToBoolean(application.DisplayAlerts, CultureInfo.InvariantCulture); }
                catch (Exception error) { return Fail("не удалось сохранить настройку предупреждений: " + ComMessage(error)); }
                try
                {
                    _workbook = _gateway.FindOpenWorkbook(_source.Workbook);
                    _sheet = _workbook.Worksheets[_source.Sheet];
                    _sourceRange = _sheet.Range[_source.Address].Cells[1, 1];
                }
                catch (Exception error) { return Fail("не удалось найти исходную ячейку: " + ComMessage(error)); }

                // The stack restores in reverse order: workbook, sheet, selection,
                // active cell, then the application settings. Each step is attempted
                // independently of the others.
                if (_originalWorkbook != null)
                    _restore.Push("активная книга", delegate { _originalWorkbook.Activate(); });
                if (_originalSheet != null)
                    _restore.Push("активный лист", delegate { _originalSheet.Activate(); });
                if (_originalSelection != null)
                    _restore.Push("выделение", delegate { _originalSelection.Select(); });
                if (_originalActiveCell != null)
                    _restore.Push("активная ячейка", delegate { _originalActiveCell.Activate(); });
                _restore.Push("предупреждения", delegate { application.DisplayAlerts = _displayAlerts; });
                _restore.Push("события", delegate { application.EnableEvents = _enableEvents; });
                _restore.Push("обновление экрана", delegate { application.ScreenUpdating = _screenUpdating; });
                return true;
            }

            // NavigateArrow fails only for a cell without visible tracer arrows, so
            // a failure means the cell has no dependent arrows of its own and it is
            // safe for the trace to add and later remove its arrows.
            private UserArrowCheck CheckUserArrows()
            {
                try
                {
                    _workbook.Activate();
                    _sheet.Activate();
                    _sourceRange.Select();
                    _sourceRange.NavigateArrow(false, 1);
                    return UserArrowCheck.Exists;
                }
                catch (Exception error)
                {
                    _lastFailure = ComMessage(error);
                    return UserArrowCheck.NoArrows;
                }
            }

            private bool CreateArrows()
            {
                dynamic application = _gateway._application;
                try
                {
                    _workbook.Activate();
                    _sheet.Activate();
                    _sourceRange.Select();
                    application.ScreenUpdating = false;
                    application.EnableEvents = false;
                    application.DisplayAlerts = false;
                    _sourceRange.ShowDependents(false);
                    return true;
                }
                catch (Exception error)
                {
                    return Fail("Excel не построил стрелки зависимых: " + ComMessage(error));
                }
            }

            private NativeArrowStep NavigateStep(int arrowNumber, int? linkNumber)
            {
                try
                {
                    _workbook.Activate();
                    _sheet.Activate();
                    _sourceRange.Select();
                    dynamic target = linkNumber.HasValue
                        ? _sourceRange.NavigateArrow(false, arrowNumber, linkNumber.Value)
                        : _sourceRange.NavigateArrow(false, arrowNumber);
                    List<NativeTraceArea> areas = CaptureAreas(target);
                    // A result that never leaves the source cell is not a dependent
                    // target; treat it as a failed navigation instead.
                    string sourceKey = _source.Workbook + "|" + _source.Sheet + "|"
                        + NativeDependentTrace.NormalizeAddress(_source.Address);
                    List<NativeTraceArea> filtered = areas.Where(area =>
                        !String.Equals(NativeDependentTrace.AreaKey(area), sourceKey, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    if (filtered.Count == 0) return NativeArrowStep.FailedStep("навигация не покинула исходную ячейку");
                    return NativeArrowStep.CapturedStep(filtered);
                }
                catch (Exception error)
                {
                    return NativeArrowStep.FailedStep(ComMessage(error));
                }
            }

            // ShowDependents(true) removes only the dependent arrows of this cell;
            // user precedent arrows are not affected and user dependent arrows
            // were excluded by the check before any arrow was created.
            private void RemoveCreatedArrows()
            {
                try
                {
                    _workbook.Activate();
                    _sheet.Activate();
                    _sourceRange.Select();
                    _sourceRange.ShowDependents(true);
                }
                catch (Exception error)
                {
                    _cleanupFailures.Add("не удалены созданные стрелки: " + ComMessage(error));
                }
            }

            private void RestoreState()
            {
                try { _restore.RestoreAll(); }
                catch (Exception error) { _cleanupFailures.Add("не восстановлено состояние: " + ComMessage(error)); }
                _cleanupFailures.AddRange(_restore.Failures);
            }

            private List<NativeTraceArea> CaptureAreas(dynamic target)
            {
                var areas = new List<NativeTraceArea>();
                if (target == null) return areas;
                try
                {
                    dynamic parts = target.Areas;
                    int count = Convert.ToInt32(parts.Count, CultureInfo.InvariantCulture);
                    for (int index = 1; index <= count; index++) CaptureSingleArea(areas, parts[index]);
                }
                catch (Exception)
                {
                    CaptureSingleArea(areas, target);
                }
                return areas;
            }

            private static void CaptureSingleArea(List<NativeTraceArea> areas, dynamic range)
            {
                dynamic sheet = range.Worksheet;
                dynamic workbook = sheet.Parent;
                areas.Add(new NativeTraceArea(
                    Convert.ToString(workbook.Name, CultureInfo.InvariantCulture),
                    Convert.ToString(sheet.Name, CultureInfo.InvariantCulture),
                    Convert.ToString(range.Address, CultureInfo.InvariantCulture)));
            }
        }
    }
}
