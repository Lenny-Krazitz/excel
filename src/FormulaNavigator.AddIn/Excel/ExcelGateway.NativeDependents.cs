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
        // Ctrl+Shift+Q entry: the direct dependents come entirely from the built-in
        // Excel trace. The reverse index is not used for this request; when the
        // trace cannot confirm its result, the captured partial result is shown
        // together with a warning instead of a silent fallback.
        private async Task<DependencyResult> TryFindNativeDependentsAsync(ExcelLocation location,
            CancellationToken cancellation, IProgress<string> progress)
        {
            var clock = Stopwatch.StartNew();
            var host = new ExcelNativeTraceHost(this, location);
            Interlocked.Increment(ref _nativeTraceActive);
            try
            {
                progress?.Report("Выполняю штатную трассировку Excel…");
                NativeTraceAttempt attempt = await NativeDependentTrace.RunAsync(host, location.Workbook,
                    location.Sheet, cancellation).ConfigureAwait(false);
                // The total request time covers the trace and the post-processing;
                // the trace stage is reported separately.
                long traceMilliseconds = clock.ElapsedMilliseconds;

                // Whatever was captured before a failure is shown to the user; it is
                // never silently dropped and never presented as confirmed.
                IReadOnlyList<NativeTraceArea> merged = NativeDependentTrace.MergeAreas(attempt.Areas);
                IReadOnlyList<NativeTraceCell> cells;
                var warnings = new List<string>();
                bool withinBudget = NativeDependentTrace.TryExpandCells(merged, NativeDependentTrace.MaxCells, out cells);
                if (!withinBudget)
                    warnings.Add("Результат усечён до " + NativeDependentTrace.MaxCells
                        + " ячеек: штатная трассировка вернула слишком большой диапазон.");
                if (attempt.Status != NativeTraceStatus.Succeeded)
                    warnings.Add(NativeDependentTrace.Decide(attempt).Warning);
                else
                    warnings.Add("Полнота результата не подтверждена: конец перечисления стрелок определяется "
                        + "по сбою NavigateArrow, который не отличает конец перечисления от сбоя существующей стрелки.");

                IReadOnlyList<ExcelLocation> detailAreas = BuildDetailAreas(merged);
                IReadOnlyList<DependentCell> placeholders = BuildPlaceholderCells(cells);
                clock.Stop();
                return new DependencyResult(placeholders, warnings.AsReadOnly(), placeholders.Count, false,
                    clock.ElapsedMilliseconds, true, traceMilliseconds, detailAreas);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception error)
            {
                return new DependencyResult(new List<DependentCell>().AsReadOnly(),
                    new List<string> { "Штатная трассировка прервана ошибкой: " + ComMessage(error) + "." }.AsReadOnly(),
                    0, false, clock.ElapsedMilliseconds, true, -1, null);
            }
            finally
            {
                Interlocked.Decrement(ref _nativeTraceActive);
            }
        }

        private static IReadOnlyList<ExcelLocation> BuildDetailAreas(IReadOnlyList<NativeTraceArea> areas)
        {
            var result = new List<ExcelLocation>();
            foreach (NativeTraceArea area in areas)
                result.Add(new ExcelLocation(area.Workbook, area.Sheet, NativeDependentTrace.NormalizeAddress(area.Address)));
            return result.AsReadOnly();
        }

        // The list is shown before the formulas and values are known, so every cell
        // first appears with an empty formula; the details task fills them in.
        private static IReadOnlyList<DependentCell> BuildPlaceholderCells(IReadOnlyList<NativeTraceCell> cells)
        {
            var result = new List<DependentCell>();
            foreach (NativeTraceCell cell in NativeDependentTrace.OrderCells(cells))
                result.Add(new DependentCell(new InspectionContext(
                    new ExcelLocation(cell.Workbook, cell.Sheet, cell.Address), String.Empty, String.Empty, false),
                    String.Empty));
            return result.AsReadOnly();
        }

        // Loads the deferred formulas and values in bounded blocks. Every area is
        // read in row blocks of at most 128 rows, each block is one batched COM
        // read (Formula2 plus Value2) dispatched separately, so Excel's thread
        // yields between blocks, cancellation is checked between them and a huge
        // area cannot occupy the thread. Every callback returns the task that
        // applies the batch, so the caller's completion means the data is visible.
        public async Task LoadNativeDetailsAsync(IReadOnlyList<ExcelLocation> areas,
            Func<IReadOnlyList<DependentDetail>, Task> onBatch, CancellationToken cancellation)
        {
            if (onBatch == null) throw new ArgumentNullException("onBatch");
            if (areas == null) return;
            var workbookCache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            int remaining = NativeDependentTrace.MaxCells;
            foreach (ExcelLocation area in areas)
            {
                ReferenceArea parsed;
                if (!ReferenceParser.TryParse(NativeDependentTrace.NormalizeAddress(area.Address), area.Workbook,
                    area.Sheet, out parsed))
                    continue;
                for (int startRow = parsed.StartRow; startRow <= parsed.EndRow; startRow += DetailBlockRows)
                {
                    cancellation.ThrowIfCancellationRequested();
                    int endRow = Math.Min(startRow + DetailBlockRows - 1, parsed.EndRow);
                    int blockCells = (endRow - startRow + 1) * (parsed.EndColumn - parsed.StartColumn + 1);
                    if (blockCells > remaining) endRow = Math.Min(endRow,
                        startRow + Math.Max(remaining / (parsed.EndColumn - parsed.StartColumn + 1), 1) - 1);
                    if (endRow < startRow) break;
                    remaining -= (endRow - startRow + 1) * (parsed.EndColumn - parsed.StartColumn + 1);
                    if (remaining < 0) remaining = 0;
                    int from = startRow, to = endRow;
                    List<DependentDetail> batch = await OnExcelThread(
                        () => ReadAreaDetails(area, parsed, from, to, workbookCache)).ConfigureAwait(false);
                    if (batch.Count > 0) await onBatch(batch).ConfigureAwait(false);
                    if (remaining == 0) return;
                }
            }
        }

        private const int DetailBlockRows = 128;

        private List<DependentDetail> ReadAreaDetails(ExcelLocation area, ReferenceArea parsed, int startRow, int endRow,
            Dictionary<string, object> workbookCache)
        {
            EnsureExcelThread();
            var batch = new List<DependentDetail>();
            try
            {
                dynamic workbook;
                object cached;
                if (workbookCache.TryGetValue(area.Workbook, out cached))
                {
                    workbook = (dynamic)cached;
                }
                else
                {
                    workbook = FindOpenWorkbook(area.Workbook);
                    workbookCache[area.Workbook] = workbook;
                }
                if (workbook == null) throw new InvalidOperationException("книга не открыта");
                dynamic sheet = workbook.Worksheets[area.Sheet];
                string blockAddress = ColumnName(parsed.StartColumn) + startRow.ToString(CultureInfo.InvariantCulture)
                    + ":" + ColumnName(parsed.EndColumn) + endRow.ToString(CultureInfo.InvariantCulture);
                dynamic range = sheet.Range[blockAddress];
                object formulas = ReadFormulas(range);
                object values;
                try { values = range.Value2; }
                catch (Exception) { values = null; }
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int column = parsed.StartColumn; column <= parsed.EndColumn; column++)
                    {
                        string address = ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
                        var cellLocation = new ExcelLocation(area.Workbook, area.Sheet, address);
                        string formula = Convert.ToString(ArrayCell(formulas, row - startRow, column - parsed.StartColumn),
                            CultureInfo.InvariantCulture) ?? String.Empty;
                        object value = ArrayCell(values, row - startRow, column - parsed.StartColumn);
                        batch.Add(new DependentDetail(cellLocation, formula, DisplayValue(value)));
                    }
                }
            }
            catch (Exception error)
            {
                // The reason is reported per cell as it happened; nothing is guessed.
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int column = parsed.StartColumn; column <= parsed.EndColumn; column++)
                    {
                        string address = ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
                        batch.Add(new DependentDetail(new ExcelLocation(area.Workbook, area.Sheet, address),
                            "не прочитано: " + ComMessage(error), String.Empty));
                    }
                }
            }
            return batch;
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

        // All tracer work runs on Excel's STA thread through the dispatcher. The
        // source sheet is activated once before the arrows are created; before a
        // navigation it is reactivated only when the previous navigation moved the
        // selection off the source sheet, which keeps cross-sheet enumeration
        // working without a fixed Activate/Select on every arrow.
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
            private bool _activated;
            private bool _leftSourceSheet;
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
                return _gateway.OnExcelThread(delegate
                {
                    RemoveCreatedArrows();
                    return true;
                });
            }

            public Task RestoreStateAsync()
            {
                return _gateway.OnExcelThread(delegate
                {
                    RestoreState();
                    return true;
                });
            }

            public void RecordCleanupFailure(string message)
            {
                _cleanupFailures.Add(message);
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

                // The stack executes in reverse order, which produces the actual
                // restore sequence: source workbook, source sheet, selection,
                // active cell, then the application settings. Application events
                // are re-enabled only after the navigation state is back.
                _restore.Push("обновление экрана", delegate { application.ScreenUpdating = _screenUpdating; });
                _restore.Push("события", delegate { application.EnableEvents = _enableEvents; });
                _restore.Push("предупреждения", delegate { application.DisplayAlerts = _displayAlerts; });
                if (_originalActiveCell != null)
                    _restore.Push("активная ячейка", delegate { _originalActiveCell.Activate(); });
                if (_originalSelection != null)
                    _restore.Push("выделение", delegate { _originalSelection.Select(); });
                if (_originalSheet != null)
                    _restore.Push("активный лист", delegate { _originalSheet.Activate(); });
                if (_originalWorkbook != null)
                    _restore.Push("активная книга", delegate { _originalWorkbook.Activate(); });
                return true;
            }

            // Preparation (activation and selection) is separate from the probe
            // itself: a failed preparation means the safety of changing arrows
            // cannot be established, and nothing is touched. A failed probe is the
            // documented "no visible tracer arrows" condition of NavigateArrow.
            private UserArrowCheck CheckUserArrows()
            {
                if (!Prepare()) return UserArrowCheck.CheckFailed;
                try
                {
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
                if (!Prepare()) return false;
                dynamic application = _gateway._application;
                try
                {
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
                    // Reactivate the source sheet only after a navigation moved the
                    // selection to another sheet; same-sheet navigation does not
                    // require a fixed Activate/Select on every arrow.
                    if (_leftSourceSheet)
                    {
                        if (!Prepare()) return NativeArrowStep.FailedStep(_lastFailure);
                        _leftSourceSheet = false;
                    }
                    dynamic target = linkNumber.HasValue
                        ? _sourceRange.NavigateArrow(false, arrowNumber, linkNumber.Value)
                        : _sourceRange.NavigateArrow(false, arrowNumber);
                    List<NativeTraceArea> areas = CaptureAreas(target);
                    string sourceKey = _source.Workbook + "|" + _source.Sheet + "|"
                        + NativeDependentTrace.NormalizeAddress(_source.Address);
                    List<NativeTraceArea> filtered = areas.Where(area =>
                        !String.Equals(NativeDependentTrace.AreaKey(area), sourceKey, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    if (filtered.Count == 0) return NativeArrowStep.FailedStep("навигация не покинула исходную ячейку");
                    _leftSourceSheet = filtered.Any(area =>
                        !String.Equals(area.Sheet, _source.Sheet, StringComparison.OrdinalIgnoreCase) ||
                        !String.Equals(area.Workbook, _source.Workbook, StringComparison.OrdinalIgnoreCase));
                    return NativeArrowStep.CapturedStep(filtered);
                }
                catch (Exception error)
                {
                    // After a failed navigation the selection state is unknown, so the
                    // next step re-prepares the source sheet.
                    _leftSourceSheet = true;
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
                    if (!_activated) Prepare();
                    _sourceRange.ShowDependents(true);
                }
                catch (Exception error)
                {
                    RecordCleanupFailure("не удалены созданные стрелки: " + ComMessage(error));
                }
            }

            private void RestoreState()
            {
                try { _restore.RestoreAll(); }
                catch (Exception error) { RecordCleanupFailure("не восстановлено состояние: " + ComMessage(error)); }
                _cleanupFailures.AddRange(_restore.Failures);
            }

            private bool Prepare()
            {
                try
                {
                    _workbook.Activate();
                    _sheet.Activate();
                    _sourceRange.Select();
                    _activated = true;
                    return true;
                }
                catch (Exception error)
                {
                    return Fail("не удалось активировать исходный лист: " + ComMessage(error));
                }
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
