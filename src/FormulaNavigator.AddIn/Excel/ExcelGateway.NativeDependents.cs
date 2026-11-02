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
        private const int DetailBlockRows = 128;

        // Ctrl+Shift+Q: a direct clone of Excel's built-in "Trace Dependents".
        // Runs synchronously on Excel's thread inside the command (macro context),
        // exactly like a VBA macro would: ShowDependents, then NavigateArrow over
        // every arrow and link, reading the target from the selection, then the
        // arrows are removed and the Excel state is restored. The formulas and
        // values of the found cells are read in batched blocks before the window
        // opens, so the window shows the complete list at once. The workbook is
        // never scanned and the reverse index is not used.
        public DependencyResult TraceDependents(ExcelLocation location)
        {
            if (location == null) throw new ArgumentNullException(nameof(location));
            EnsureExcelThread();
            var total = Stopwatch.StartNew();
            var host = new ExcelNativeTraceHost(this, location);
            Interlocked.Increment(ref _nativeTraceActive);
            NativeTraceAttempt attempt;
            try
            {
                // Every host method runs inline and returns a completed task, so the
                // shared enumerator completes synchronously without a thread switch.
                attempt = NativeDependentTrace.RunAsync(host, location.Workbook, location.Sheet, CancellationToken.None)
                    .GetAwaiter().GetResult();
            }
            catch (Exception error)
            {
                Interlocked.Decrement(ref _nativeTraceActive);
                return new DependencyResult(new List<DependentCell>().AsReadOnly(),
                    new List<string> { "Штатная трассировка Excel прервана ошибкой: " + ComMessage(error) }.AsReadOnly(),
                    0, false, total.ElapsedMilliseconds, true, total.ElapsedMilliseconds, null);
            }
            Interlocked.Decrement(ref _nativeTraceActive);
            long traceMilliseconds = total.ElapsedMilliseconds;

            var warnings = new List<string>();
            if (attempt.Status != NativeTraceStatus.Succeeded)
                warnings.Add(NativeDependentTrace.Decide(attempt).Warning);
            IReadOnlyList<NativeTraceArea> merged = NativeDependentTrace.MergeAreas(attempt.Areas);
            IReadOnlyList<NativeTraceCell> expanded;
            if (!NativeDependentTrace.TryExpandCells(merged, NativeDependentTrace.MaxCells, out expanded))
                warnings.Add("Результат усечён до " + NativeDependentTrace.MaxCells
                    + " ячеек: штатная трассировка вернула слишком большой диапазон.");

            var cells = ReadDependentCells(merged);
            total.Stop();
            return new DependencyResult(cells.AsReadOnly(), warnings.AsReadOnly(), cells.Count, false,
                total.ElapsedMilliseconds, true, traceMilliseconds, null);
        }

        // Reads formulas and values area by area in blocks of 128 rows: one
        // Formula2 and one Value2 call per block. A failed block is reported per
        // cell with the actual reason.
        private List<DependentCell> ReadDependentCells(IReadOnlyList<NativeTraceArea> areas)
        {
            var cells = new List<DependentCell>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var workbooks = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            int remaining = NativeDependentTrace.MaxCells;
            foreach (NativeTraceArea area in areas)
            {
                ReferenceArea parsed;
                if (!ReferenceParser.TryParse(area.Address, area.Workbook, area.Sheet, out parsed)) continue;
                int width = parsed.EndColumn - parsed.StartColumn + 1;
                for (int startRow = parsed.StartRow; startRow <= parsed.EndRow && remaining > 0; startRow += DetailBlockRows)
                {
                    int endRow = Math.Min(startRow + DetailBlockRows - 1, parsed.EndRow);
                    endRow = Math.Min(endRow, startRow + Math.Max(remaining / width, 1) - 1);
                    ReadBlock(area, parsed, startRow, endRow, workbooks, seen, cells);
                    remaining = NativeDependentTrace.MaxCells - cells.Count;
                }
                if (remaining <= 0) break;
            }
            return cells;
        }

        private void ReadBlock(NativeTraceArea area, ReferenceArea parsed, int startRow, int endRow,
            Dictionary<string, object> workbooks, HashSet<string> seen, List<DependentCell> cells)
        {
            object formulas = null;
            object values = null;
            string failure = null;
            try
            {
                object cached;
                dynamic workbook;
                if (workbooks.TryGetValue(area.Workbook, out cached)) workbook = cached;
                else
                {
                    workbook = FindOpenWorkbook(area.Workbook);
                    workbooks[area.Workbook] = workbook;
                }
                if (workbook == null) throw new InvalidOperationException("книга не открыта");
                dynamic sheet = workbook.Worksheets[area.Sheet];
                string address = ColumnName(parsed.StartColumn) + startRow.ToString(CultureInfo.InvariantCulture)
                    + ":" + ColumnName(parsed.EndColumn) + endRow.ToString(CultureInfo.InvariantCulture);
                dynamic range = sheet.Range[address];
                formulas = ReadFormulas(range);
                try { values = range.Value2; }
                catch (Exception) { values = null; }
            }
            catch (Exception error) { failure = "не прочитано: " + ComMessage(error); }

            for (int row = startRow; row <= endRow; row++)
            {
                for (int column = parsed.StartColumn; column <= parsed.EndColumn; column++)
                {
                    string address = ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
                    var location = new ExcelLocation(area.Workbook, area.Sheet, address);
                    if (!seen.Add(location.Key)) continue;
                    if (failure != null)
                    {
                        cells.Add(new DependentCell(new InspectionContext(location, String.Empty, String.Empty, false), failure));
                        continue;
                    }
                    string formula = Convert.ToString(ArrayCell(formulas, row - startRow, column - parsed.StartColumn),
                        CultureInfo.InvariantCulture) ?? String.Empty;
                    object value = ArrayCell(values, row - startRow, column - parsed.StartColumn);
                    cells.Add(new DependentCell(new InspectionContext(location, formula, DisplayValue(value), IsFormula(formula)),
                        String.Empty));
                }
            }
        }

        // Excel returns either a scalar (single-cell range) or a 1-based 2D array.
        private static object ArrayCell(object values, int rowOffset, int columnOffset)
        {
            Array array = values as Array;
            if (array == null) return values;
            if (array.Rank != 2) return null;
            int row = rowOffset + array.GetLowerBound(0);
            int column = columnOffset + array.GetLowerBound(1);
            if (row > array.GetUpperBound(0) || column > array.GetUpperBound(1)) return null;
            return array.GetValue(row, column);
        }

        // All work runs inline on Excel's thread. Every method returns a completed
        // task. NavigateArrow moves the selection to its target; the target is read
        // from Application.Selection, and a selection that stayed on the source
        // cell is the end of the arrow or link.
        private sealed class ExcelNativeTraceHost : INativeTraceHost
        {
            private readonly ExcelGateway _gateway;
            private readonly ExcelLocation _source;
            private readonly string _sourceKey;
            private readonly StateRestoreStack _restore = new StateRestoreStack();
            private readonly List<string> _cleanupFailures = new List<string>();
            private dynamic _workbook;
            private dynamic _sheet;
            private dynamic _sourceRange;
            private bool _offSourceSheet;
            private string _lastFailure;

            public ExcelNativeTraceHost(ExcelGateway gateway, ExcelLocation source)
            {
                _gateway = gateway;
                _source = source;
                _sourceKey = Key(source.Workbook, source.Sheet, source.Address);
            }

            public IReadOnlyList<string> CleanupFailures { get { return _cleanupFailures.AsReadOnly(); } }
            public string LastFailure { get { return _lastFailure; } }

            public Task<bool> TrySaveStateAsync() { return Task.FromResult(SaveState()); }
            public Task<UserArrowCheck> CheckUserArrowsAsync() { return Task.FromResult(CheckUserArrows()); }
            public Task<bool> TryCreateArrowsAsync() { return Task.FromResult(CreateArrows()); }
            public Task<NativeArrowStep> NavigateAsync(int arrowNumber, int? linkNumber)
            {
                return Task.FromResult(Navigate(arrowNumber, linkNumber ?? 1));
            }
            public Task RemoveCreatedArrowsAsync() { RemoveCreatedArrows(); return Task.FromResult(true); }
            public Task RestoreStateAsync() { RestoreState(); return Task.FromResult(true); }
            public void RecordCleanupFailure(string message) { _cleanupFailures.Add(message); }

            private bool Fail(string message)
            {
                _lastFailure = message;
                return false;
            }

            private bool SaveState()
            {
                dynamic application = _gateway._application;
                dynamic originalWorkbook, originalSheet, originalSelection, originalActiveCell;
                bool screenUpdating, enableEvents, displayAlerts;
                try
                {
                    originalWorkbook = application.ActiveWorkbook;
                    originalSheet = application.ActiveSheet;
                    originalSelection = application.Selection;
                    originalActiveCell = application.ActiveCell;
                    screenUpdating = Convert.ToBoolean(application.ScreenUpdating, CultureInfo.InvariantCulture);
                    enableEvents = Convert.ToBoolean(application.EnableEvents, CultureInfo.InvariantCulture);
                    displayAlerts = Convert.ToBoolean(application.DisplayAlerts, CultureInfo.InvariantCulture);
                }
                catch (Exception error) { return Fail("не удалось сохранить состояние Excel: " + ComMessage(error)); }
                try
                {
                    _workbook = _gateway.FindOpenWorkbook(_source.Workbook);
                    if (_workbook == null) return Fail("книга '" + _source.Workbook + "' не открыта");
                    _sheet = _workbook.Worksheets[_source.Sheet];
                    _sourceRange = _sheet.Range[_source.Address].Cells[1, 1];
                }
                catch (Exception error) { return Fail("не удалось найти исходную ячейку: " + ComMessage(error)); }

                // Executed in reverse: workbook, sheet, selection, active cell, then
                // the application settings.
                _restore.Push("обновление экрана", delegate { application.ScreenUpdating = screenUpdating; });
                _restore.Push("события", delegate { application.EnableEvents = enableEvents; });
                _restore.Push("предупреждения", delegate { application.DisplayAlerts = displayAlerts; });
                if (originalActiveCell != null) _restore.Push("активная ячейка", delegate { originalActiveCell.Activate(); });
                if (originalSelection != null) _restore.Push("выделение", delegate { originalSelection.Select(); });
                if (originalSheet != null) _restore.Push("активный лист", delegate { originalSheet.Activate(); });
                if (originalWorkbook != null) _restore.Push("активная книга", delegate { originalWorkbook.Activate(); });

                return true;
            }

            // Existing arrows of the source cell: following arrow 1 moves the
            // selection away. Without arrows the selection stays or Excel raises
            // an error. Either way the selection is returned to the source.
            private UserArrowCheck CheckUserArrows()
            {
                try
                {
                    dynamic application = _gateway._application;
                    application.ScreenUpdating = false;
                    application.EnableEvents = false;
                    application.DisplayAlerts = false;
                }
                catch (Exception error)
                {
                    Fail("не удалось отключить обновление экрана: " + ComMessage(error));
                    return UserArrowCheck.CheckFailed;
                }
                if (!GoToSource(true)) return UserArrowCheck.CheckFailed;
                NativeArrowStep probe = Navigate(1, 1);
                if (!GoToSource(true)) return UserArrowCheck.CheckFailed;
                return probe.Captured ? UserArrowCheck.Exists : UserArrowCheck.NoArrows;
            }

            private bool CreateArrows()
            {
                if (!GoToSource(true)) return false;
                try
                {
                    _sourceRange.ShowDependents(false);
                    return true;
                }
                catch (Exception error) { return Fail("Excel не построил стрелки зависимых: " + ComMessage(error)); }
            }

            private NativeArrowStep Navigate(int arrowNumber, int linkNumber)
            {
                // NavigateArrow works relative to the selected source cell, as in
                // the built-in command. Selecting a cell with screen updating off
                // is cheap; the sheet is reactivated only after leaving it.
                if (!GoToSource(_offSourceSheet)) return NativeArrowStep.FailedStep(_lastFailure);
                try
                {
                    _sourceRange.NavigateArrow(false, arrowNumber, linkNumber);
                }
                catch (Exception error)
                {
                    _offSourceSheet = true;
                    _lastFailure = ComMessage(error);
                    return NativeArrowStep.FailedStep(_lastFailure);
                }
                try
                {
                    dynamic selection = _gateway._application.Selection;
                    List<NativeTraceArea> areas = CaptureAreas(selection);
                    if (areas.Count == 0) return NativeArrowStep.EndStep();
                    _offSourceSheet = areas.Any(area =>
                        !String.Equals(area.Sheet, _source.Sheet, StringComparison.OrdinalIgnoreCase) ||
                        !String.Equals(area.Workbook, _source.Workbook, StringComparison.OrdinalIgnoreCase));
                    bool stayed = areas.Count == 1 && String.Equals(Key(areas[0].Workbook, areas[0].Sheet, areas[0].Address),
                        _sourceKey, StringComparison.OrdinalIgnoreCase);
                    return stayed ? NativeArrowStep.EndStep() : NativeArrowStep.CapturedStep(areas);
                }
                catch (Exception error)
                {
                    _offSourceSheet = true;
                    _lastFailure = "не удалось прочитать цель стрелки: " + ComMessage(error);
                    return NativeArrowStep.FailedStep(_lastFailure);
                }
            }

            private void RemoveCreatedArrows()
            {
                try
                {
                    GoToSource(_offSourceSheet);
                    _sourceRange.ShowDependents(true);
                }
                catch (Exception error) { RecordCleanupFailure("не удалены созданные стрелки: " + ComMessage(error)); }
            }

            private void RestoreState()
            {
                _restore.RestoreAll();
                _cleanupFailures.AddRange(_restore.Failures);
            }

            private bool GoToSource(bool activate)
            {
                try
                {
                    if (activate)
                    {
                        _workbook.Activate();
                        _sheet.Activate();
                        _offSourceSheet = false;
                    }
                    _sourceRange.Select();
                    return true;
                }
                catch (Exception error) { return Fail("не удалось выделить исходную ячейку: " + ComMessage(error)); }
            }

            private static List<NativeTraceArea> CaptureAreas(dynamic range)
            {
                var areas = new List<NativeTraceArea>();
                if (range == null) return areas;
                dynamic parts = range.Areas;
                int count = Convert.ToInt32(parts.Count, CultureInfo.InvariantCulture);
                for (int index = 1; index <= count; index++)
                {
                    dynamic part = parts[index];
                    dynamic sheet = part.Worksheet;
                    dynamic workbook = sheet.Parent;
                    areas.Add(new NativeTraceArea(
                        Convert.ToString(workbook.Name, CultureInfo.InvariantCulture),
                        Convert.ToString(sheet.Name, CultureInfo.InvariantCulture),
                        NativeDependentTrace.NormalizeAddress(Convert.ToString(part.Address, CultureInfo.InvariantCulture))));
                }
                return areas;
            }

            private static string Key(string workbook, string sheet, string address)
            {
                return workbook + "|" + sheet + "|" + NativeDependentTrace.NormalizeAddress(address);
            }
        }
    }
}
