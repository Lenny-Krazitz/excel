using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace FormulaNavigator.Core
{
    // One area of the range returned by Range.NavigateArrow, captured as plain data.
    // The range may consist of several areas located on other sheets or in other
    // open workbooks; every area is recorded with its own workbook and sheet name.
    public sealed class NativeTraceArea
    {
        public NativeTraceArea(string workbook, string sheet, string address)
        {
            if (String.IsNullOrWhiteSpace(workbook)) throw new ArgumentException("A workbook name is required.", "workbook");
            if (String.IsNullOrWhiteSpace(sheet)) throw new ArgumentException("A worksheet name is required.", "sheet");
            if (String.IsNullOrWhiteSpace(address)) throw new ArgumentException("An A1 address is required.", "address");
            Workbook = workbook;
            Sheet = sheet;
            Address = address;
        }

        public string Workbook { get; private set; }
        public string Sheet { get; private set; }
        public string Address { get; private set; }
    }

    // A single dependent cell produced by expanding captured areas.
    public sealed class NativeTraceCell
    {
        public NativeTraceCell(string workbook, string sheet, string address, int row, int column)
        {
            Workbook = workbook;
            Sheet = sheet;
            Address = address;
            Row = row;
            Column = column;
        }

        public string Workbook { get; private set; }
        public string Sheet { get; private set; }
        public string Address { get; private set; }
        public int Row { get; private set; }
        public int Column { get; private set; }
    }

    public enum NativeTraceStatus
    {
        // Enumeration finished and every step, including state restore, succeeded.
        Succeeded,
        // The workbook/sheet/selection snapshot could not be read; nothing was changed.
        SetupFailed,
        // The check for existing user arrows failed; nothing was changed.
        UserArrowCheckFailed,
        // The cell already has dependent tracer arrows; nothing was changed.
        ExistingUserArrows,
        // Excel could not create the tracer arrows; state was restored.
        ShowDependentsFailed,
        // The first arrow could not be navigated: "no arrows" (no dependents) and a
        // COM failure are indistinguishable, so the result cannot be confirmed.
        FirstArrowFailed,
        // A later arrow or link failed and even a repeated probe could not confirm
        // the end of enumeration; a partial result must not be reported as complete.
        EnumerationFailed,
        // A protective limit was reached; the result may be incomplete.
        LimitReached,
        // Removing the created arrows or restoring the Excel state failed; the
        // result is not trusted even if enumeration itself succeeded.
        RestoreFailed
    }

    public sealed class NativeTraceAttempt
    {
        private readonly List<NativeTraceArea> _areas = new List<NativeTraceArea>();

        public NativeTraceStatus Status { get; private set; }
        public string Failure { get; private set; }

        public IReadOnlyList<NativeTraceArea> Areas { get { return _areas.AsReadOnly(); } }

        public void Succeed()
        {
            Status = NativeTraceStatus.Succeeded;
            Failure = null;
        }

        public void Fail(NativeTraceStatus status, string failure)
        {
            Status = status;
            Failure = failure;
        }

        public void AddAreas(IEnumerable<NativeTraceArea> areas)
        {
            if (areas == null) return;
            foreach (NativeTraceArea area in areas) _areas.Add(area);
        }
    }

    // Outcome of one NavigateArrow call. The COM adapter cannot distinguish the
    // documented "no visible tracer arrows" error from any other failure, so it
    // reports only Captured/Failed and the enumerator classifies the failure.
    public sealed class NativeArrowStep
    {
        private NativeArrowStep() { }

        public bool Captured { get; private set; }
        public IReadOnlyList<NativeTraceArea> Areas { get; private set; }
        public string Failure { get; private set; }

        public static NativeArrowStep CapturedStep(IReadOnlyList<NativeTraceArea> areas)
        {
            return new NativeArrowStep { Captured = true, Areas = areas ?? new List<NativeTraceArea>().AsReadOnly() };
        }

        public static NativeArrowStep FailedStep(string failure)
        {
            return new NativeArrowStep { Captured = false, Failure = failure };
        }
    }

    public enum UserArrowCheck
    {
        // NavigateArrow failed, which per its contract means the cell has no
        // visible tracer arrows, so creating and later removing arrows is safe.
        NoArrows,
        // The cell already has dependent tracer arrows; they must not be touched.
        Exists,
        // The probe itself failed; safety of arrow removal cannot be established.
        CheckFailed
    }

    // Host operations performed on Excel's STA thread by the gateway. Every method
    // reports success or a failure reason; the trace never proceeds after a failed
    // step and always cleans up what it has changed.
    public interface INativeTraceHost
    {
        Task<bool> TrySaveStateAsync();
        Task<UserArrowCheck> CheckUserArrowsAsync();
        Task<bool> TryCreateArrowsAsync();
        Task<NativeArrowStep> NavigateAsync(int arrowNumber, int? linkNumber);
        // Records its own failures; called only after the user-arrow check
        // confirmed that the cell has no arrows of its own.
        Task RemoveCreatedArrowsAsync();
        // Records its own failures through the shared restore stack.
        Task RestoreStateAsync();
        // Records a cleanup failure that happened outside the two methods above,
        // for example when the dispatch of a cleanup step itself failed.
        void RecordCleanupFailure(string message);
        IReadOnlyList<string> CleanupFailures { get; }
        string LastFailure { get; }
    }

    public sealed class NativeTraceDecision
    {
        public NativeTraceDecision(bool useNative, string warning)
        {
            UseNative = useNative;
            Warning = warning;
        }

        public bool UseNative { get; private set; }
        public string Warning { get; private set; }
    }

    public static class NativeDependentTrace
    {
        // Protective caps. Excel does not expose the number of tracer arrows, so
        // enumeration ends on the first failing NavigateArrow; the caps only bound
        // a pathological workbook.
        public const int MaxArrows = 16384;
        public const int MaxLinks = 16384;
        // Upper bound on the number of dependent cells expanded from the captured
        // areas before the add-in reads their formulas and values.
        public const int MaxCells = 16384;

        public static async Task<NativeTraceAttempt> RunAsync(INativeTraceHost host, string sourceWorkbook,
            string sourceSheet, CancellationToken cancellation)
        {
            if (host == null) throw new ArgumentNullException("host");
            var attempt = new NativeTraceAttempt();
            if (!await host.TrySaveStateAsync().ConfigureAwait(false))
            {
                attempt.Fail(NativeTraceStatus.SetupFailed, host.LastFailure);
                return attempt;
            }
            // Nothing has been changed yet; from here on the finally block must
            // remove created arrows and restore the saved state. Every step of the
            // cleanup runs even when a previous step failed, so a late error can
            // never skip the state restoration.
            bool mayRemoveArrows = false;
            try
            {
                UserArrowCheck check = await host.CheckUserArrowsAsync().ConfigureAwait(false);
                if (check == UserArrowCheck.CheckFailed)
                {
                    attempt.Fail(NativeTraceStatus.UserArrowCheckFailed, host.LastFailure);
                    return attempt;
                }
                if (check == UserArrowCheck.Exists)
                {
                    attempt.Fail(NativeTraceStatus.ExistingUserArrows,
                        "У исходной ячейки уже есть пользовательские стрелки зависимых; штатная трассировка отменена, чтобы не удалить их.");
                    return attempt;
                }
                mayRemoveArrows = true;
                if (!await host.TryCreateArrowsAsync().ConfigureAwait(false))
                {
                    attempt.Fail(NativeTraceStatus.ShowDependentsFailed, host.LastFailure);
                    return attempt;
                }
                // Every captured target is recorded immediately, so a later failure
                // still leaves the partial result available to the user.
                NativeTraceStatus status = await EnumerateAsync(host, sourceWorkbook, sourceSheet, attempt, cancellation)
                    .ConfigureAwait(false);
                if (status != NativeTraceStatus.Succeeded)
                {
                    attempt.Fail(status, FailureMessage(status, host.LastFailure));
                    return attempt;
                }
                attempt.Succeed();
            }
            finally
            {
                if (mayRemoveArrows)
                {
                    try { await host.RemoveCreatedArrowsAsync().ConfigureAwait(false); }
                    catch (Exception error) { host.RecordCleanupFailure("не удалены созданные стрелки: " + error.Message); }
                }
                try { await host.RestoreStateAsync().ConfigureAwait(false); }
                catch (Exception error) { host.RecordCleanupFailure("не восстановлено состояние: " + error.Message); }
            }
            if (attempt.Status == NativeTraceStatus.Succeeded && host.CleanupFailures.Count != 0)
                attempt.Fail(NativeTraceStatus.RestoreFailed, String.Join("; ", host.CleanupFailures.ToArray()));
            return attempt;
        }

        private static async Task<NativeTraceStatus> EnumerateAsync(INativeTraceHost host, string sourceWorkbook,
            string sourceSheet, NativeTraceAttempt attempt, CancellationToken cancellation)
        {
            for (int arrow = 1; arrow <= MaxArrows; arrow++)
            {
                cancellation.ThrowIfCancellationRequested();
                NativeArrowStep step = await host.NavigateAsync(arrow, null).ConfigureAwait(false);
                if (!step.Captured)
                {
                    if (arrow == 1) return NativeTraceStatus.FirstArrowFailed;
                    // Confirm the trace is still healthy: the previous arrow must
                    // still navigate. If it does, a repeated failure of the current
                    // arrow is the documented "no such arrow" condition. This does
                    // not prove that a later arrow does not exist: the failure of an
                    // existing arrow is indistinguishable from its absence, so the
                    // completeness of the result stays unconfirmed.
                    NativeArrowStep recheck = await host.NavigateAsync(arrow - 1, null).ConfigureAwait(false);
                    if (!recheck.Captured) return NativeTraceStatus.EnumerationFailed;
                    NativeArrowStep retry = await host.NavigateAsync(arrow, null).ConfigureAwait(false);
                    if (!retry.Captured) break;
                    step = retry;
                }
                attempt.AddAreas(step.Areas);
                if (arrow == MaxArrows) return NativeTraceStatus.LimitReached;
                // Excel combines every target that leaves the source sheet — other
                // sheets of the same workbook as well as other workbooks — into a
                // single off-sheet arrow and enumerates the individual targets
                // through LinkNumber. Same-sheet targets get their own arrows.
                bool offSheet = step.Areas.Any(area =>
                    !String.Equals(area.Sheet, sourceSheet, StringComparison.OrdinalIgnoreCase));
                if (!offSheet) continue;
                for (int link = 2; link <= MaxLinks; link++)
                {
                    cancellation.ThrowIfCancellationRequested();
                    NativeArrowStep linkStep = await host.NavigateAsync(arrow, link).ConfigureAwait(false);
                    if (!linkStep.Captured)
                    {
                        NativeArrowStep recheck = await host.NavigateAsync(arrow, link - 1).ConfigureAwait(false);
                        if (!recheck.Captured) return NativeTraceStatus.EnumerationFailed;
                        NativeArrowStep retry = await host.NavigateAsync(arrow, link).ConfigureAwait(false);
                        if (!retry.Captured) break;
                        linkStep = retry;
                    }
                    attempt.AddAreas(linkStep.Areas);
                    if (link == MaxLinks) return NativeTraceStatus.LimitReached;
                }
            }
            return NativeTraceStatus.Succeeded;
        }

        public static NativeTraceDecision Decide(NativeTraceAttempt attempt)
        {
            if (attempt == null) throw new ArgumentNullException("attempt");
            if (attempt.Status == NativeTraceStatus.Succeeded)
                return new NativeTraceDecision(true, null);
            return new NativeTraceDecision(false, FailureMessage(attempt.Status, attempt.Failure));
        }

        private static string FailureMessage(NativeTraceStatus status, string detail)
        {
            string message;
            switch (status)
            {
                case NativeTraceStatus.SetupFailed:
                    message = "Не удалось сохранить исходное состояние Excel; штатная трассировка не выполнялась.";
                    break;
                case NativeTraceStatus.UserArrowCheckFailed:
                    message = "Не удалось проверить отсутствие пользовательских стрелок; штатная трассировка не выполнялась.";
                    break;
                case NativeTraceStatus.ExistingUserArrows:
                    message = "У исходной ячейки есть пользовательские стрелки зависимых; они не изменялись.";
                    break;
                case NativeTraceStatus.ShowDependentsFailed:
                    message = "Excel не построил стрелки зависимых (возможна защита листа или неподдерживаемый случай).";
                    break;
                case NativeTraceStatus.FirstArrowFailed:
                    message = "Не удалось подтвердить наличие или отсутствие зависимых: сбой первой стрелки неотличим от их отсутствия.";
                    break;
                case NativeTraceStatus.EnumerationFailed:
                    message = "Штатная трассировка прервана ошибкой; результат был бы неполным.";
                    break;
                case NativeTraceStatus.LimitReached:
                    message = "Превышен защитный предел стрелок или внешних ссылок; результат был бы неполным.";
                    break;
                case NativeTraceStatus.RestoreFailed:
                    message = "Не удалось полностью восстановить состояние Excel после трассировки.";
                    break;
                default:
                    message = "Штатная трассировка не завершена.";
                    break;
            }
            if (!String.IsNullOrEmpty(detail)) message += " " + detail;
            return message;
        }

        // Strips absolute markers and collapses degenerate ranges: "$A$1:$A$1" and
        // "$A$1" both normalize to "A1".
        public static string NormalizeAddress(string address)
        {
            string text = (address ?? String.Empty).Replace("$", String.Empty);
            int colon = text.IndexOf(':');
            if (colon < 0) return text;
            string first = text.Substring(0, colon);
            string last = text.Substring(colon + 1);
            return String.Equals(first, last, StringComparison.OrdinalIgnoreCase) ? first : first + ":" + last;
        }

        // Deduplicates captured areas by workbook, sheet and normalized address.
        public static IReadOnlyList<NativeTraceArea> MergeAreas(IEnumerable<NativeTraceArea> areas)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var merged = new List<NativeTraceArea>();
            if (areas == null) return merged.AsReadOnly();
            foreach (NativeTraceArea area in areas)
            {
                string key = AreaKey(area);
                if (!seen.Add(key)) continue;
                merged.Add(new NativeTraceArea(area.Workbook, area.Sheet, NormalizeAddress(area.Address)));
            }
            return merged.AsReadOnly();
        }

        // Expands areas into individual cells, deduplicating overlapping areas at
        // the cell level. Returns false when the result would exceed the budget:
        // the caller must fall back to the index instead of truncating.
        public static bool TryExpandCells(IEnumerable<NativeTraceArea> areas, int budget, out IReadOnlyList<NativeTraceCell> cells)
        {
            var result = new List<NativeTraceCell>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            cells = result.AsReadOnly();
            if (areas == null) return true;
            foreach (NativeTraceArea area in areas)
            {
                ReferenceArea parsed;
                if (!ReferenceParser.TryParse(NormalizeAddress(area.Address), area.Workbook, area.Sheet, out parsed))
                    return false;
                long count = (long)(parsed.EndRow - parsed.StartRow + 1) * (parsed.EndColumn - parsed.StartColumn + 1);
                if (result.Count + count > budget) return false;
                for (int row = parsed.StartRow; row <= parsed.EndRow; row++)
                {
                    for (int column = parsed.StartColumn; column <= parsed.EndColumn; column++)
                    {
                        string address = ColumnName(column) + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        if (!seen.Add(area.Workbook + "|" + area.Sheet + "|" + address)) continue;
                        result.Add(new NativeTraceCell(area.Workbook, area.Sheet, address, row, column));
                    }
                }
            }
            return true;
        }

        // Orders expanded cells deterministically: workbook, sheet, row, column.
        public static IReadOnlyList<NativeTraceCell> OrderCells(IEnumerable<NativeTraceCell> cells)
        {
            if (cells == null) return new List<NativeTraceCell>().AsReadOnly();
            return cells.OrderBy(c => c.Workbook, StringComparer.OrdinalIgnoreCase)
                .ThenBy(c => c.Sheet, StringComparer.OrdinalIgnoreCase)
                .ThenBy(c => c.Row).ThenBy(c => c.Column).ToList().AsReadOnly();
        }

        public static string AreaKey(NativeTraceArea area)
        {
            return area.Workbook + "|" + area.Sheet + "|" + NormalizeAddress(area.Address);
        }

        private static string ColumnName(int column)
        {
            string result = String.Empty;
            while (column > 0)
            {
                column--;
                result = (char)('A' + column % 26) + result;
                column /= 26;
            }
            return result;
        }
    }

    // Records independent restore actions and executes them in reverse order.
    // Every action is attempted even if a previous one failed, and failures are
    // collected so the caller can refuse to report a silent success.
    public sealed class StateRestoreStack
    {
        private readonly List<Step> _steps = new List<Step>();
        private readonly List<string> _failures = new List<string>();

        public IReadOnlyList<string> Failures { get { return _failures.AsReadOnly(); } }

        public void Push(string name, Action restore)
        {
            if (String.IsNullOrEmpty(name)) throw new ArgumentException("A step name is required.", "name");
            if (restore == null) throw new ArgumentNullException("restore");
            _steps.Add(new Step(name, restore));
        }

        public void RestoreAll()
        {
            for (int i = _steps.Count - 1; i >= 0; i--) _steps[i].Run(_failures);
        }

        private sealed class Step
        {
            private readonly string _name;
            private readonly Action _restore;

            public Step(string name, Action restore)
            {
                _name = name;
                _restore = restore;
            }

            public void Run(List<string> failures)
            {
                try { _restore(); }
                catch (Exception) { failures.Add("не восстановлено: " + _name); }
            }
        }
    }
}
