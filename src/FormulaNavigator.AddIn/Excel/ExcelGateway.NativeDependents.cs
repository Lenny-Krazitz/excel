using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FormulaNavigator.AddIn.Excel
{
    public sealed partial class ExcelGateway
    {
        public async Task<string> CompareNativeDependentsAsync(ExcelLocation location, CancellationToken cancellation)
        {
            if (location == null) throw new ArgumentNullException(nameof(location));
            NativeTraceResult native = await OnExcelThread(() => TraceNativeDependents(location, cancellation))
                .ConfigureAwait(false);
            cancellation.ThrowIfCancellationRequested();

            var currentClock = Stopwatch.StartNew();
            DependencyResult current = await FindDependentsAsync(location, cancellation, null).ConfigureAwait(false);
            currentClock.Stop();

            var nativeKeys = new HashSet<string>(native.Locations.Select(LocationKey), StringComparer.OrdinalIgnoreCase);
            var currentKeys = new HashSet<string>(current.Cells.Select(cell => LocationKey(cell.Context.Location)),
                StringComparer.OrdinalIgnoreCase);
            string[] missingFromNative = currentKeys.Except(nativeKeys, StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToArray();
            string[] nativeOnly = nativeKeys.Except(currentKeys, StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToArray();

            var report = new StringBuilder();
            report.AppendLine("Эксперимент ShowDependents + NavigateArrow");
            report.AppendLine("Источник: " + location);
            report.AppendLine();
            report.AppendLine("Штатная трассировка: " + native.Locations.Count + " уникальных ячеек, "
                + native.ElapsedMilliseconds + " мс.");
            report.AppendLine("Текущий индекс: " + current.Cells.Count + " уникальных ячеек, "
                + currentClock.ElapsedMilliseconds + " мс (внутреннее измерение " + current.ElapsedMilliseconds
                + " мс, индекс " + (current.IndexReused ? "переиспользован" : "подготовлен") + ").");
            if (!String.IsNullOrEmpty(native.Warning)) report.AppendLine("Штатная трассировка: " + native.Warning);
            if (native.Inconclusive)
            {
                report.AppendLine();
                report.AppendLine("Результат НЕПОЛНЫЙ/НЕПРОВЕРЕННЫЙ; рабочий Ctrl+Shift+Q не изменён.");
            }
            else
            {
                report.AppendLine();
                report.AppendLine(missingFromNative.Length == 0 && nativeOnly.Length == 0
                    ? "Наборы результатов совпали. Это один замер; повторите на книгах с именами, таблицами, динамическими ссылками и межлистовыми зависимостями."
                    : "Наборы результатов различаются; штатную трассировку нельзя включать как основной путь.");
                AppendDifference(report, "Не найдены штатной трассировкой", missingFromNative);
                AppendDifference(report, "Найдены только штатной трассировкой", nativeOnly);
            }
            report.AppendLine();
            report.AppendLine("Прототип не меняет Ctrl+Shift+Q, не перестраивает индекс и не повышает версию.");
            return report.ToString();
        }

        private NativeTraceResult TraceNativeDependents(ExcelLocation location, CancellationToken cancellation)
        {
            EnsureExcelThread();
            var clock = Stopwatch.StartNew();
            dynamic originalWorkbook = null;
            dynamic originalSheet = null;
            dynamic originalSelection = null;
            bool screenUpdating = true;
            bool enableEvents = true;
            bool displayAlerts = true;
            bool arrowsCreated = false;
            var locations = new Dictionary<string, ExcelLocation>(StringComparer.OrdinalIgnoreCase);
            string warning = null;
            bool inconclusive = false;

            try
            {
                try { originalWorkbook = _application.ActiveWorkbook; } catch { }
                try { originalSheet = _application.ActiveSheet; } catch { }
                try { originalSelection = _application.Selection; } catch { }
                try { screenUpdating = Convert.ToBoolean(_application.ScreenUpdating, CultureInfo.InvariantCulture); } catch { }
                try { enableEvents = Convert.ToBoolean(_application.EnableEvents, CultureInfo.InvariantCulture); } catch { }
                try { displayAlerts = Convert.ToBoolean(_application.DisplayAlerts, CultureInfo.InvariantCulture); } catch { }

                _application.ScreenUpdating = false;
                _application.EnableEvents = false;
                _application.DisplayAlerts = false;

                dynamic workbook = FindOpenWorkbook(location.Workbook);
                dynamic sheet = workbook.Worksheets[location.Sheet];
                dynamic source = sheet.Range[location.Address].Cells[1, 1];
                workbook.Activate();
                sheet.Activate();
                source.Select();

                if (HasVisibleDependentArrow(source, workbook, sheet))
                {
                    inconclusive = true;
                    warning = "У исходной ячейки уже есть пользовательские стрелки зависимых. Эксперимент отменён, чтобы не удалить их.";
                    return new NativeTraceResult(locations.Values.ToList().AsReadOnly(), warning, true,
                        clock.ElapsedMilliseconds);
                }

                arrowsCreated = true;
                try { source.ShowDependents(false); }
                catch (Exception error)
                {
                    inconclusive = true;
                    warning = "Excel не построил стрелки (это может означать отсутствие зависимых либо неподдерживаемый/защищённый лист): "
                        + ComMessage(error);
                    return new NativeTraceResult(locations.Values.ToList().AsReadOnly(), warning, true,
                        clock.ElapsedMilliseconds);
                }

                const int maximumArrows = 16384;
                const int maximumLinks = 16384;
                for (int arrow = 1; arrow <= maximumArrows; arrow++)
                {
                    cancellation.ThrowIfCancellationRequested();
                    List<ExcelLocation> first;
                    if (!TryNavigateDependent(source, workbook, sheet, arrow, null, out first)) break;
                    AddLocations(locations, first);

                    bool leavesSourceSheet = first.Any(item =>
                        !String.Equals(item.Workbook, location.Workbook, StringComparison.OrdinalIgnoreCase)
                        || !String.Equals(item.Sheet, location.Sheet, StringComparison.OrdinalIgnoreCase));
                    if (!leavesSourceSheet) continue;

                    for (int link = 1; link <= maximumLinks; link++)
                    {
                        cancellation.ThrowIfCancellationRequested();
                        List<ExcelLocation> linked;
                        if (!TryNavigateDependent(source, workbook, sheet, arrow, link, out linked)) break;
                        AddLocations(locations, linked);
                        if (link == maximumLinks)
                        {
                            inconclusive = true;
                            warning = "Превышен защитный предел внешних ссылок; результат может быть неполным.";
                        }
                    }
                    if (arrow == maximumArrows)
                    {
                        inconclusive = true;
                        warning = "Превышен защитный предел стрелок; результат может быть неполным.";
                    }
                }
            }
            finally
            {
                if (arrowsCreated)
                {
                    try
                    {
                        dynamic workbook = FindOpenWorkbook(location.Workbook);
                        dynamic sheet = workbook.Worksheets[location.Sheet];
                        dynamic source = sheet.Range[location.Address].Cells[1, 1];
                        workbook.Activate();
                        sheet.Activate();
                        source.Select();
                        source.ShowDependents(true);
                    }
                    catch { }
                }
                try { if (originalWorkbook != null) originalWorkbook.Activate(); } catch { }
                try { if (originalSheet != null) originalSheet.Activate(); } catch { }
                try { if (originalSelection != null) originalSelection.Select(); } catch { }
                try { _application.DisplayAlerts = displayAlerts; } catch { }
                try { _application.EnableEvents = enableEvents; } catch { }
                try { _application.ScreenUpdating = screenUpdating; } catch { }
            }

            clock.Stop();
            return new NativeTraceResult(locations.Values.OrderBy(item => item.Workbook)
                .ThenBy(item => item.Sheet).ThenBy(item => item.Address).ToList().AsReadOnly(), warning,
                inconclusive, clock.ElapsedMilliseconds);
        }

        private static bool HasVisibleDependentArrow(dynamic source, dynamic workbook, dynamic sheet)
        {
            try
            {
                source.NavigateArrow(false, 1);
                return true;
            }
            catch { return false; }
            finally
            {
                try { workbook.Activate(); } catch { }
                try { sheet.Activate(); } catch { }
                try { source.Select(); } catch { }
            }
        }

        private bool TryNavigateDependent(dynamic source, dynamic workbook, dynamic sourceSheet, int arrow,
            int? link, out List<ExcelLocation> locations)
        {
            locations = new List<ExcelLocation>();
            try
            {
                workbook.Activate();
                sourceSheet.Activate();
                source.Select();
                dynamic target = link.HasValue
                    ? source.NavigateArrow(false, arrow, link.Value)
                    : source.NavigateArrow(false, arrow);
                CaptureRangeLocations(target, locations);
                return locations.Count != 0;
            }
            catch { return false; }
            finally
            {
                try { workbook.Activate(); } catch { }
                try { sourceSheet.Activate(); } catch { }
                try { source.Select(); } catch { }
            }
        }

        private static void CaptureRangeLocations(dynamic target, List<ExcelLocation> locations)
        {
            if (target == null) return;
            try
            {
                dynamic areas = target.Areas;
                int count = Convert.ToInt32(areas.Count, CultureInfo.InvariantCulture);
                for (int index = 1; index <= count; index++) CaptureSingleRange(areas[index], locations);
            }
            catch { CaptureSingleRange(target, locations); }
        }

        private static void CaptureSingleRange(dynamic range, List<ExcelLocation> locations)
        {
            dynamic sheet = range.Worksheet;
            dynamic workbook = sheet.Parent;
            string workbookName = Convert.ToString(workbook.Name, CultureInfo.InvariantCulture);
            string sheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);
            string address = Convert.ToString(range.Address, CultureInfo.InvariantCulture);
            locations.Add(new ExcelLocation(workbookName, sheetName, address));
        }

        private static void AddLocations(IDictionary<string, ExcelLocation> destination,
            IEnumerable<ExcelLocation> locations)
        {
            foreach (ExcelLocation location in locations) destination[LocationKey(location)] = location;
        }

        private static string LocationKey(ExcelLocation location)
        {
            return location.Workbook + "|" + location.Sheet + "|" + location.Address.Replace("$", String.Empty);
        }

        private static void AppendDifference(StringBuilder report, string title, IReadOnlyList<string> values)
        {
            if (values.Count == 0) return;
            report.AppendLine(title + " (" + values.Count + "):");
            int shown = Math.Min(values.Count, 20);
            for (int i = 0; i < shown; i++) report.AppendLine("  " + values[i]);
            if (values.Count > shown) report.AppendLine("  … ещё " + (values.Count - shown));
        }

        private sealed class NativeTraceResult
        {
            public NativeTraceResult(IReadOnlyList<ExcelLocation> locations, string warning, bool inconclusive,
                long elapsedMilliseconds)
            {
                Locations = locations;
                Warning = warning;
                Inconclusive = inconclusive;
                ElapsedMilliseconds = elapsedMilliseconds;
            }

            public readonly IReadOnlyList<ExcelLocation> Locations;
            public readonly string Warning;
            public readonly bool Inconclusive;
            public readonly long ElapsedMilliseconds;
        }
    }
}
