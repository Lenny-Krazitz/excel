using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using FormulaNavigator.Core;

namespace FormulaNavigator.AddIn.Excel
{
    // All methods which touch a COM object run through the Dispatcher captured on Excel's STA.
    // The dependency walk deliberately never uses Task.Run.
    public sealed class ExcelGateway
    {
        private const int XlCellTypeFormulas = -4123;
        private const int MaxCellsPerRead = 4096;
        private readonly dynamic _application;
        private readonly Dispatcher _excelDispatcher;
        private readonly int _excelThreadId;

        public ExcelGateway(object excelApplication)
        {
            if (excelApplication == null) throw new ArgumentNullException("excelApplication");
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("ExcelGateway must be created on Excel's STA thread.");
            _application = excelApplication;
            _excelThreadId = Thread.CurrentThread.ManagedThreadId;
            _excelDispatcher = Dispatcher.CurrentDispatcher;
        }

        public InspectionContext CaptureActiveCell()
        {
            EnsureExcelThread();
            dynamic cell = _application.ActiveCell;
            if (cell == null) throw new InvalidOperationException("Excel does not have an active cell.");
            return CaptureContext(cell);
        }

        public InspectionContext Capture(ExcelLocation location)
        {
            EnsureExcelThread();
            if (location == null) throw new ArgumentNullException("location");
            dynamic workbook = FindOpenWorkbook(location.Workbook);
            dynamic worksheet = workbook.Worksheets[location.Sheet];
            dynamic range = worksheet.Range[location.Address];
            return CaptureContext(range.Cells[1, 1]);
        }

        public NodeInspection Inspect(InspectionContext context, FormulaNode node)
        {
            EnsureExcelThread();
            if (context == null) throw new ArgumentNullException("context");
            if (node == null) throw new ArgumentNullException("node");

            try
            {
                dynamic worksheet = GetWorksheet(context.Location);
                Resolution resolution;
                string resolutionNote;
                if (TryResolveNodeAsRange(context, worksheet, node, out resolution, out resolutionNote))
                {
                    return new NodeInspection(DisplayResolvedRange(resolution), resolutionNote, ToLocation(resolution));
                }

                object value;
                string evaluationNote;
                if (TryEvaluateSafe(context, worksheet, node, out value, out evaluationNote))
                    return new NodeInspection(DisplayValue(value), evaluationNote, null);

                return new NodeInspection(String.Empty,
                    String.IsNullOrEmpty(evaluationNote) ? resolutionNote : evaluationNote, null);
            }
            catch (Exception ex)
            {
                return new NodeInspection(String.Empty, "Could not inspect this expression: " + ComMessage(ex), null);
            }
        }

        public void Navigate(ExcelLocation location)
        {
            EnsureExcelThread();
            if (location == null) throw new ArgumentNullException("location");
            dynamic workbook = FindOpenWorkbook(location.Workbook);
            dynamic worksheet = workbook.Worksheets[location.Sheet];
            dynamic range = worksheet.Range[location.Address];
            workbook.Activate();
            worksheet.Activate();
            range.Select();
        }

        public IntPtr GetWindowHandle()
        {
            EnsureExcelThread();
            try
            {
                return new IntPtr(Convert.ToInt64(_application.Hwnd, CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Excel's window handle is unavailable.", ex);
            }
        }

        public Task<DependencyResult> FindDependentsAsync(ExcelLocation location, CancellationToken cancellation,
            IProgress<string> progress)
        {
            if (location == null) throw new ArgumentNullException("location");
            var completion = new TaskCompletionSource<DependencyResult>();
            Action start = delegate
            {
                if (cancellation.IsCancellationRequested)
                {
                    completion.TrySetCanceled();
                    return;
                }

                try
                {
                    var state = new DependencyScan(this, location, cancellation, progress, completion);
                    state.Start();
                }
                catch (OperationCanceledException)
                {
                    completion.TrySetCanceled();
                }
                catch (Exception ex)
                {
                    completion.TrySetException(ex);
                }
            };

            try
            {
                if (_excelDispatcher.CheckAccess() && Thread.CurrentThread.ManagedThreadId == _excelThreadId)
                    start();
                else
                    _excelDispatcher.BeginInvoke(start);
            }
            catch (Exception ex)
            {
                completion.TrySetException(new InvalidOperationException(
                    "Excel's STA dispatcher is unavailable for dependency scanning.", ex));
            }

            return completion.Task;
        }

        private void EnsureExcelThread()
        {
            if (Thread.CurrentThread.ManagedThreadId != _excelThreadId || !_excelDispatcher.CheckAccess())
                throw new InvalidOperationException("Excel COM must be accessed from the Excel STA thread that created ExcelGateway.");
        }

        private InspectionContext CaptureContext(dynamic range)
        {
            range = range.Cells[1, 1];
            dynamic sheet = range.Worksheet;
            dynamic workbook = sheet.Parent;
            string formula = ReadFormula(range);
            object value = range.Value2;
            var location = new ExcelLocation(Convert.ToString(workbook.Name, CultureInfo.InvariantCulture),
                Convert.ToString(sheet.Name, CultureInfo.InvariantCulture), GetAddress(range));
            return new InspectionContext(location, formula, DisplayValue(value), IsFormula(formula));
        }

        private dynamic FindOpenWorkbook(string workbookName)
        {
            dynamic workbooks = _application.Workbooks;
            int count = Convert.ToInt32(workbooks.Count, CultureInfo.InvariantCulture);
            for (int i = 1; i <= count; i++)
            {
                dynamic workbook = workbooks[i];
                if (String.Equals(Convert.ToString(workbook.Name, CultureInfo.InvariantCulture), workbookName,
                    StringComparison.OrdinalIgnoreCase)) return workbook;
            }
            throw new InvalidOperationException("Workbook '" + workbookName + "' is not open.");
        }

        private dynamic GetWorksheet(ExcelLocation location)
        {
            dynamic workbook = FindOpenWorkbook(location.Workbook);
            return workbook.Worksheets[location.Sheet];
        }

        private static string ReadFormula(dynamic range)
        {
            object formula;
            try { formula = range.Formula2; }
            catch { formula = range.Formula; }
            return formula as string ?? String.Empty;
        }

        private static object ReadFormulas(dynamic range)
        {
            try { return range.Formula2; }
            catch { return range.Formula; }
        }

        private static bool IsFormula(string formula)
        {
            return !String.IsNullOrEmpty(formula) && formula[0] == '=';
        }

        private static string GetAddress(dynamic range)
        {
            // Excel's default Range.Address is an A1 address local to its worksheet.
            return Convert.ToString(range.Address, CultureInfo.InvariantCulture);
        }

        private static string ComMessage(Exception exception)
        {
            return String.IsNullOrWhiteSpace(exception.Message) ? exception.GetType().Name : exception.Message;
        }

        private static string DisplayValue(object value)
        {
            if (value == null) return String.Empty;
            string excelError;
            if (TryDisplayExcelError(value, out excelError)) return excelError;
            Array array = value as Array;
            if (array != null)
                return String.Format(CultureInfo.InvariantCulture, "{0} × {1} cells", array.GetLength(0), array.GetLength(1));
            if (value is DateTime) return ((DateTime)value).ToString(CultureInfo.CurrentCulture);
            return Convert.ToString(value, CultureInfo.CurrentCulture) ?? String.Empty;
        }

        private static bool TryDisplayExcelError(object value, out string display)
        {
            display = null;
            var wrapper = value as ErrorWrapper;
            if (wrapper != null) return TryExcelErrorCode(wrapper.ErrorCode, out display);
            if (value is int && (int)value < 0) return TryExcelErrorCode((int)value, out display);
            return false;
        }

        private static bool TryExcelErrorCode(int code, out string display)
        {
            // Excel may expose VT_ERROR as either a raw SCODE or an HRESULT whose low word is the SCODE.
            switch (code & 0xffff)
            {
                case 2000: display = "#NULL!"; return true;
                case 2007: display = "#DIV/0!"; return true;
                case 2015: display = "#VALUE!"; return true;
                case 2023: display = "#REF!"; return true;
                case 2029: display = "#NAME?"; return true;
                case 2036: display = "#NUM!"; return true;
                case 2042: display = "#N/A"; return true;
                case 2043: display = "#GETTING_DATA"; return true;
                case 2045: display = "#SPILL!"; return true;
                case 2046: display = "#CONNECT!"; return true;
                case 2047: display = "#BLOCKED!"; return true;
                case 2048: display = "#UNKNOWN!"; return true;
                case 2049: display = "#FIELD!"; return true;
                case 2050: display = "#CALC!"; return true;
                default: display = null; return false;
            }
        }

        private static ExcelLocation ToLocation(Resolution resolution)
        {
            return new ExcelLocation(resolution.Area.Workbook, resolution.Area.Sheet, GetA1Address(resolution.Area));
        }

        private static string GetA1Address(ReferenceArea area)
        {
            string first = ColumnName(area.StartColumn) + area.StartRow.ToString(CultureInfo.InvariantCulture);
            string last = ColumnName(area.EndColumn) + area.EndRow.ToString(CultureInfo.InvariantCulture);
            return first == last ? first : first + ":" + last;
        }

        private static string ColumnName(int column)
        {
            string result = String.Empty;
            while (column > 0)
            {
                column--;
                result = (char)('A' + (column % 26)) + result;
                column /= 26;
            }
            return result;
        }

        private string DisplayResolvedRange(Resolution resolution)
        {
            long count = (long)(resolution.Area.EndRow - resolution.Area.StartRow + 1)
                * (resolution.Area.EndColumn - resolution.Area.StartColumn + 1);
            if (count != 1) return "Range " + GetA1Address(resolution.Area) + " (" + count.ToString(CultureInfo.InvariantCulture) + " cells)";
            try { return DisplayValue(resolution.Range.Value2); }
            catch (Exception ex) { return "Value unavailable: " + ComMessage(ex); }
        }

        private bool TryResolveNodeAsRange(InspectionContext context, dynamic worksheet, FormulaNode node,
            out Resolution resolution, out string note)
        {
            resolution = null;
            note = String.Empty;
            if (node.Kind == FormulaNodeKind.Reference)
                return TryResolveReference(context, worksheet, node.Text, false, out resolution, out note);
            if (node.Kind == FormulaNodeKind.Name)
                return TryResolveName(context, worksheet, node.Text, out resolution, out note);
            if (node.Kind == FormulaNodeKind.Binary && node.Operator == ":")
                return TryResolveRangeOperator(context, worksheet, node, out resolution, out note);
            if (node.Kind == FormulaNodeKind.Unary && node.Operator == "#")
                return TryResolveSpill(context, worksheet, node, out resolution, out note);
            if (node.Kind != FormulaNodeKind.Function) return false;

            string name = (node.Name ?? String.Empty).ToUpperInvariant();
            if (name == "OFFSET") return TryResolveOffset(context, worksheet, node, out resolution, out note);
            if (name == "INDIRECT") return TryResolveIndirect(context, worksheet, node, out resolution, out note);
            return false;
        }

        private bool TryResolveSpill(InspectionContext context, dynamic worksheet, FormulaNode node, out Resolution resolution,
            out string note)
        {
            resolution = null;
            note = String.Empty;
            if (node.Children == null || node.Children.Count != 1)
            {
                note = "The spill operator is incomplete.";
                return false;
            }
            Resolution anchor;
            if (!TryResolveNodeAsRange(context, worksheet, node.Children[0], out anchor, out note)) return false;
            if (anchor.Area.StartRow != anchor.Area.EndRow || anchor.Area.StartColumn != anchor.Area.EndColumn)
            {
                note = "The spill operator requires a single-cell anchor.";
                return false;
            }
            try
            {
                dynamic spilled = anchor.Range.SpillingToRange;
                dynamic targetSheet = spilled.Worksheet;
                dynamic targetBook = targetSheet.Parent;
                string workbook = Convert.ToString(targetBook.Name, CultureInfo.InvariantCulture);
                if (!String.Equals(workbook, context.Location.Workbook, StringComparison.OrdinalIgnoreCase))
                {
                    note = "The spill range resolves to an external workbook and was not followed.";
                    return false;
                }
                int row = Convert.ToInt32(spilled.Row, CultureInfo.InvariantCulture);
                int column = Convert.ToInt32(spilled.Column, CultureInfo.InvariantCulture);
                int rows = Convert.ToInt32(spilled.Rows.Count, CultureInfo.InvariantCulture);
                int columns = Convert.ToInt32(spilled.Columns.Count, CultureInfo.InvariantCulture);
                resolution = new Resolution(new ReferenceArea(workbook,
                    Convert.ToString(targetSheet.Name, CultureInfo.InvariantCulture), row, row + rows - 1,
                    column, column + columns - 1), spilled);
                return true;
            }
            catch (Exception ex)
            {
                note = "The spilled range is unavailable: " + ComMessage(ex);
                return false;
            }
        }

        private bool TryResolveRangeOperator(InspectionContext context, dynamic worksheet, FormulaNode node,
            out Resolution resolution, out string note)
        {
            resolution = null;
            note = String.Empty;
            if (node.Children == null || node.Children.Count != 2)
            {
                note = "The range operator is incomplete.";
                return false;
            }
            Resolution first, second;
            if (!TryResolveNodeAsRange(context, (object)worksheet, node.Children[0], out first, out note)
                || !TryResolveNodeAsRange(context, (object)worksheet, node.Children[1], out second, out note)) return false;
            // In Sheet2!A1:B5 Excel applies Sheet2 to the unqualified right endpoint.
            // The parser retains the two endpoint nodes, so carry that qualifier here.
            if ((!String.Equals(first.Area.Workbook, second.Area.Workbook, StringComparison.OrdinalIgnoreCase)
                    || !String.Equals(first.Area.Sheet, second.Area.Sheet, StringComparison.OrdinalIgnoreCase))
                && node.Children[0].Kind == FormulaNodeKind.Reference
                && node.Children[1].Kind == FormulaNodeKind.Reference
                && FindQualifierBang(node.Children[0].Text) >= 0
                && FindQualifierBang(node.Children[1].Text) < 0)
            {
                var qualifiedContext = new InspectionContext(
                    new ExcelLocation(first.Area.Workbook, first.Area.Sheet, context.Location.Address),
                    context.Formula, context.DisplayValue, context.HasFormula);
                if (!TryResolveReference(qualifiedContext, worksheet, node.Children[1].Text, false, out second, out note)) return false;
            }
            if (!String.Equals(first.Area.Workbook, second.Area.Workbook, StringComparison.OrdinalIgnoreCase)
                || !String.Equals(first.Area.Sheet, second.Area.Sheet, StringComparison.OrdinalIgnoreCase))
            {
                note = "A range cannot span two worksheets in this inspector.";
                return false;
            }
            var area = new ReferenceArea(first.Area.Workbook, first.Area.Sheet,
                Math.Min(first.Area.StartRow, second.Area.StartRow), Math.Max(first.Area.EndRow, second.Area.EndRow),
                Math.Min(first.Area.StartColumn, second.Area.StartColumn), Math.Max(first.Area.EndColumn, second.Area.EndColumn));
            try
            {
                dynamic targetSheet = GetWorksheet(new ExcelLocation(area.Workbook, area.Sheet, "A1"));
                resolution = new Resolution(area, targetSheet.Range[GetA1Address(area)]);
                return true;
            }
            catch (Exception ex)
            {
                note = "The range is unavailable: " + ComMessage(ex);
                return false;
            }
        }

        private bool TryResolveReference(InspectionContext context, dynamic worksheet, string text, bool r1c1,
            out Resolution resolution, out string note)
        {
            resolution = null;
            note = String.Empty;
            ReferenceArea area;
            if (r1c1)
            {
                if (!TryParseR1C1(text, context.Location, out area))
                {
                    note = "The R1C1 reference could not be resolved from the source cell.";
                    return false;
                }
            }
            else if (!ReferenceParser.TryParse(text, context.Location.Workbook, context.Location.Sheet, out area))
            {
                note = "This reference form is not supported.";
                return false;
            }

            if (!String.Equals(area.Workbook, context.Location.Workbook, StringComparison.OrdinalIgnoreCase))
            {
                note = "This reference points to an external or closed workbook and was not resolved.";
                return false;
            }
            try
            {
                dynamic targetSheet = GetWorksheet(new ExcelLocation(area.Workbook, area.Sheet, "A1"));
                dynamic range = targetSheet.Range[GetA1Address(area)];
                resolution = new Resolution(area, range);
                return true;
            }
            catch (Exception ex)
            {
                note = "The referenced sheet or range is unavailable: " + ComMessage(ex);
                return false;
            }
        }

        private bool TryResolveName(InspectionContext context, dynamic worksheet, string text, out Resolution resolution,
            out string note)
        {
            resolution = null;
            note = String.Empty;
            ReferenceArea r1c1Area;
            if (TryParseR1C1(text, context.Location, out r1c1Area))
            {
                return TryResolveReference(context, worksheet, text, true, out resolution, out note);
            }
            if (IsRelativeDefinedName(worksheet, text))
            {
                note = "Relative defined names are not resolved because their meaning can depend on Excel's active cell.";
                return false;
            }
            try
            {
                // Range is Excel's own name/table resolver; no text rewriting or guessed scope is used.
                dynamic range = worksheet.Range[text];
                dynamic targetSheet = range.Worksheet;
                dynamic targetBook = targetSheet.Parent;
                string bookName = Convert.ToString(targetBook.Name, CultureInfo.InvariantCulture);
                if (!String.Equals(bookName, context.Location.Workbook, StringComparison.OrdinalIgnoreCase))
                {
                    note = "The name resolves to an external workbook and was not followed.";
                    return false;
                }
                int row = Convert.ToInt32(range.Row, CultureInfo.InvariantCulture);
                int column = Convert.ToInt32(range.Column, CultureInfo.InvariantCulture);
                int rows = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
                int columns = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
                resolution = new Resolution(new ReferenceArea(bookName,
                    Convert.ToString(targetSheet.Name, CultureInfo.InvariantCulture), row, row + rows - 1,
                    column, column + columns - 1), range);
                return true;
            }
            catch (Exception ex)
            {
                note = "This name or structured reference does not resolve to a worksheet range: " + ComMessage(ex);
                return false;
            }
        }

        private static bool IsRelativeDefinedName(dynamic worksheet, string text)
        {
            // Table selectors are resolved by Worksheet.Range and are not defined names.
            if (String.IsNullOrWhiteSpace(text) || text.IndexOf('[') >= 0 || text.IndexOf(']') >= 0) return false;
            try
            {
                dynamic name = worksheet.Names[text];
                if (IsRelativeRefersTo(Convert.ToString(name.RefersToR1C1, CultureInfo.InvariantCulture))) return true;
            }
            catch { }
            try
            {
                dynamic name = worksheet.Parent.Names[text];
                if (IsRelativeRefersTo(Convert.ToString(name.RefersToR1C1, CultureInfo.InvariantCulture))) return true;
            }
            catch { }
            return false;
        }

        private static bool IsRelativeRefersTo(string refersToR1C1)
        {
            return !String.IsNullOrEmpty(refersToR1C1)
                && (refersToR1C1.IndexOf("R[", StringComparison.OrdinalIgnoreCase) >= 0
                    || refersToR1C1.IndexOf("C[", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private bool TryResolveOffset(InspectionContext context, dynamic worksheet, FormulaNode node, out Resolution resolution,
            out string note)
        {
            resolution = null;
            note = String.Empty;
            if (node.Children == null || node.Children.Count < 3 || node.Children.Count > 5 || HasMissingArgument(node.Children))
            {
                note = "OFFSET requires three to five non-empty arguments.";
                return false;
            }
            Resolution baseResolution;
            if (!TryResolveNodeAsRange(context, worksheet, node.Children[0], out baseResolution, out note))
            {
                if (String.IsNullOrEmpty(note)) note = "The OFFSET base is not a resolvable range.";
                return false;
            }
            double rowsValue, columnsValue;
            if (!TryEvaluateNumber(context, (object)worksheet, node.Children[1], out rowsValue, out note)
                || !TryEvaluateNumber(context, (object)worksheet, node.Children[2], out columnsValue, out note))
            {
                note = "OFFSET row and column arguments must evaluate safely to numbers. " + note;
                return false;
            }
            int rowsOffset, columnsOffset;
            if (!TryIntegral(rowsValue, out rowsOffset) || !TryIntegral(columnsValue, out columnsOffset))
            {
                note = "OFFSET row and column arguments must be whole numbers.";
                return false;
            }
            int height = baseResolution.Area.EndRow - baseResolution.Area.StartRow + 1;
            int width = baseResolution.Area.EndColumn - baseResolution.Area.StartColumn + 1;
            if (node.Children.Count >= 4 && !TryDimension(context, worksheet, node.Children[3], out height, out note)) return false;
            if (node.Children.Count >= 5 && !TryDimension(context, worksheet, node.Children[4], out width, out note)) return false;
            int row = baseResolution.Area.StartRow + rowsOffset;
            int column = baseResolution.Area.StartColumn + columnsOffset;
            if (row < 1 || column < 1 || height < 1 || width < 1
                || row + height - 1 > ReferenceParser.MaxRows || column + width - 1 > ReferenceParser.MaxColumns)
            {
                note = "OFFSET resolves outside the worksheet.";
                return false;
            }
            var area = new ReferenceArea(baseResolution.Area.Workbook, baseResolution.Area.Sheet, row, row + height - 1,
                column, column + width - 1);
            try
            {
                dynamic targetSheet = GetWorksheet(new ExcelLocation(area.Workbook, area.Sheet, "A1"));
                resolution = new Resolution(area, targetSheet.Range[GetA1Address(area)]);
                return true;
            }
            catch (Exception ex)
            {
                note = "OFFSET target is unavailable: " + ComMessage(ex);
                return false;
            }
        }

        private bool TryResolveIndirect(InspectionContext context, dynamic worksheet, FormulaNode node, out Resolution resolution,
            out string note)
        {
            resolution = null;
            note = String.Empty;
            if (node.Children == null || node.Children.Count == 0 || node.Children.Count > 2 || HasMissingArgument(node.Children))
            {
                note = "INDIRECT requires one or two non-empty arguments.";
                return false;
            }
            object referenceText;
            if (!TryEvaluateSafe(context, worksheet, node.Children[0], out referenceText, out note)
                || !(referenceText is string))
            {
                note = "INDIRECT text could not be evaluated safely.";
                return false;
            }
            bool r1c1 = false;
            if (node.Children.Count > 1)
            {
                object a1;
                if (!TryEvaluateSafe(context, worksheet, node.Children[1], out a1, out note)
                    || !TryBoolean(a1, out r1c1))
                {
                    note = "INDIRECT's A1/R1C1 argument could not be evaluated safely.";
                    return false;
                }
                r1c1 = !r1c1;
            }
            if (TryResolveReference(context, worksheet, (string)referenceText, r1c1, out resolution, out note)) return true;
            if (!r1c1)
            {
                string referenceNote = note;
                ReferenceArea parsedArea;
                // Do not retry an A1-looking (including external) reference through Range,
                // because that could make Excel follow a workbook reference.
                if (!ReferenceParser.TryParse((string)referenceText, context.Location.Workbook, context.Location.Sheet, out parsedArea)
                    && TryResolveName(context, worksheet, (string)referenceText, out resolution, out note)) return true;
                if (String.IsNullOrEmpty(note)) note = referenceNote;
            }
            return false;
        }

        private static bool HasMissingArgument(IReadOnlyList<FormulaNode> children)
        {
            foreach (FormulaNode child in children)
                if (child == null || child.Kind == FormulaNodeKind.MissingArgument) return true;
            return false;
        }

        private static bool TryIntegral(double value, out int result)
        {
            result = 0;
            if (Double.IsNaN(value) || Double.IsInfinity(value) || Math.Floor(value) != value
                || value > Int32.MaxValue || value < Int32.MinValue) return false;
            result = (int)value;
            return true;
        }

        private bool TryDimension(InspectionContext context, dynamic worksheet, FormulaNode node, out int result, out string note)
        {
            result = 0;
            double value;
            if (!TryEvaluateNumber(context, worksheet, node, out value, out note)) return false;
            if (!TryIntegral(value, out result) || result < 1)
            {
                note = "OFFSET height and width must be positive whole numbers.";
                return false;
            }
            return true;
        }

        private bool TryEvaluateNumber(InspectionContext context, dynamic worksheet, FormulaNode node, out double value,
            out string note)
        {
            value = 0;
            object raw;
            if (!TryEvaluateSafe(context, worksheet, node, out raw, out note)) return false;
            string error;
            if (TryDisplayExcelError(raw, out error)) { note = error; return false; }
            try
            {
                value = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return !Double.IsNaN(value) && !Double.IsInfinity(value);
            }
            catch
            {
                note = "The expression is not numeric.";
                return false;
            }
        }

        private bool TryEvaluateSafe(InspectionContext context, dynamic worksheet, FormulaNode node, out object value,
            out string note)
        {
            value = null;
            string rendered;
            var renderer = new SafeExpressionRenderer(this, context, worksheet);
            if (!renderer.TryRender(node, out rendered, out note)) return false;
            string expression = "=" + rendered;
            if (expression.Length > 255)
            {
                note = "This safe expression is longer than Excel's 255-character Evaluate limit.";
                return false;
            }
            try
            {
                value = worksheet.Evaluate(expression);
                note = String.Empty;
                return true;
            }
            catch (Exception ex)
            {
                note = "Safe evaluation failed: " + ComMessage(ex);
                return false;
            }
        }

        private static bool TryBoolean(object value, out bool result)
        {
            result = false;
            if (value is bool) { result = (bool)value; return true; }
            string error;
            if (TryDisplayExcelError(value, out error)) return false;
            try
            {
                double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                result = number != 0;
                return true;
            }
            catch { }
            string text = value as string;
            if (Boolean.TryParse(text, out result)) return true;
            return false;
        }

        // This renderer is deliberately narrower than Excel's formula language.  It turns
        // parsed, validated nodes into an expression whose references are concrete A1 ranges,
        // then Worksheet.Evaluate supplies Excel's own coercion and error semantics.
        private sealed class SafeExpressionRenderer
        {
            private static readonly HashSet<string> SafeFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ABS", "AND", "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "CEILING", "CEILING.MATH", "CHOOSE",
                "CHOOSECOLS", "CHOOSEROWS", "COLUMN", "COLUMNS", "CONCAT", "CONCATENATE", "COUNT",
                "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "DATE", "DATEDIF", "DATEVALUE", "DAY",
                "DAYS", "DAYS360", "EDATE", "EOMONTH", "EVEN", "EXACT", "EXP", "FILTER", "FIND",
                "FLOOR", "FLOOR.MATH", "HLOOKUP", "HOUR", "IF", "IFERROR", "IFNA", "INDEX", "INT",
                "IPMT", "IRR", "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISFORMULA", "ISLOGICAL",
                "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISTEXT", "LARGE", "LEFT", "LEN", "LN",
                "LOG", "LOG10", "LOWER", "MATCH", "MAX", "MAXIFS", "MEDIAN", "MID", "MIN", "MINIFS",
                "MINUTE", "MIRR", "MOD", "MONTH", "NETWORKDAYS", "NOT", "NPER", "NPV", "ODD", "OR",
                "PMT", "POWER", "PPMT", "PRODUCT", "PROPER", "PV", "QUOTIENT", "RANK",
                "RATE", "REPLACE", "RIGHT", "ROUND", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS", "SEARCH",
                "SECOND", "SIGN", "SMALL", "SORT", "SQRT", "SUBSTITUTE", "SUM", "SUMIF", "SUMIFS",
                "TEXT", "TEXTJOIN", "TIME", "TIMEVALUE", "TRIM", "TRUNC", "UPPER", "VALUE", "VLOOKUP",
                "WEEKDAY", "WEEKNUM", "WORKDAY", "XLOOKUP", "XMATCH", "XOR", "YEAR"
            };

            private static readonly HashSet<string> ErrorLiterals = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA",
                "#SPILL!", "#CONNECT!", "#BLOCKED!", "#UNKNOWN!", "#FIELD!", "#CALC!"
            };

            private readonly ExcelGateway _gateway;
            private readonly InspectionContext _context;
            private readonly dynamic _worksheet;

            public SafeExpressionRenderer(ExcelGateway gateway, InspectionContext context, dynamic worksheet)
            {
                _gateway = gateway;
                _context = context;
                _worksheet = worksheet;
            }

            public bool TryRender(FormulaNode node, out string rendered, out string note)
            {
                return Render(node, false, out rendered, out note);
            }

            private bool Render(FormulaNode node, bool functionArgument, out string rendered, out string note)
            {
                rendered = null;
                note = String.Empty;
                if (node == null) { note = "The expression is missing."; return false; }
                if (node.Kind == FormulaNodeKind.MissingArgument)
                {
                    if (functionArgument) { rendered = String.Empty; return true; }
                    note = "A missing argument cannot be evaluated outside a function call.";
                    return false;
                }
                if (node.Kind == FormulaNodeKind.Reference || node.Kind == FormulaNodeKind.Name
                    || node.Kind == FormulaNodeKind.Binary && node.Operator == ":"
                    || node.Kind == FormulaNodeKind.Unary && node.Operator == "#")
                {
                    Resolution range;
                    if (!_gateway.TryResolveNodeAsRange(_context, _worksheet, node, out range, out note)) return false;
                    return RenderRange(range, out rendered, out note);
                }
                if (node.Kind == FormulaNodeKind.Literal)
                {
                    if (!IsSafeLiteral(node.Text))
                    {
                        note = "This literal is not safe to render for evaluation.";
                        return false;
                    }
                    rendered = node.Text;
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Group)
                {
                    if (node.Children == null || node.Children.Count != 1) { note = "The grouped expression is incomplete."; return false; }
                    string child;
                    if (!Render(node.Children[0], false, out child, out note)) return false;
                    rendered = "(" + child + ")";
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Unary)
                {
                    if (node.Children == null || node.Children.Count != 1 || !IsSafeUnary(node.Operator))
                    {
                        note = "This unary operation is not safe to evaluate.";
                        return false;
                    }
                    string child;
                    if (!Render(node.Children[0], false, out child, out note)) return false;
                    rendered = node.Operator == "%" ? "(" + child + ")%" : node.Operator + "(" + child + ")";
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Binary)
                {
                    if (node.Children == null || node.Children.Count != 2 || !IsSafeBinary(node.Operator))
                    {
                        note = "This binary operation is not safe to evaluate.";
                        return false;
                    }
                    string left, right;
                    if (!Render(node.Children[0], false, out left, out note)
                        || !Render(node.Children[1], false, out right, out note)) return false;
                    rendered = "(" + left + ")" + node.Operator + "(" + right + ")";
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Function)
                    return RenderFunction(node, out rendered, out note);

                note = "This expression kind is not safe to evaluate.";
                return false;
            }

            private bool RenderFunction(FormulaNode node, out string rendered, out string note)
            {
                rendered = null;
                note = String.Empty;
                string name = (node.Name ?? String.Empty).ToUpperInvariant();
                if (name == "OFFSET" || name == "INDIRECT")
                {
                    Resolution range;
                    bool resolved = name == "OFFSET"
                        ? _gateway.TryResolveOffset(_context, _worksheet, node, out range, out note)
                        : _gateway.TryResolveIndirect(_context, _worksheet, node, out range, out note);
                    return resolved && RenderRange(range, out rendered, out note);
                }
                if (name == "LET" || name == "LAMBDA")
                {
                    note = name + " uses lexical scope and is not evaluated by the inspector.";
                    return false;
                }
                if (!SafeFunctions.Contains(name))
                {
                    note = "Function '" + (node.Name ?? node.Text) + "' is not in the safe evaluation allowlist.";
                    return false;
                }
                if ((name == "ROW" || name == "COLUMN") && (node.Children == null || node.Children.Count == 0))
                {
                    note = name + " without a reference depends on Excel's active cell and is unavailable.";
                    return false;
                }
                var arguments = new List<string>();
                if (node.Children != null)
                {
                    foreach (FormulaNode child in node.Children)
                    {
                        string argument;
                        if (!Render(child, true, out argument, out note)) return false;
                        arguments.Add(argument);
                    }
                }
                rendered = name + "(" + String.Join(",", arguments) + ")";
                return true;
            }

            private bool RenderRange(Resolution range, out string rendered, out string note)
            {
                rendered = null;
                note = String.Empty;
                if (!String.Equals(range.Area.Workbook, _context.Location.Workbook, StringComparison.OrdinalIgnoreCase))
                {
                    note = "The resolved range is outside the captured workbook.";
                    return false;
                }
                string address = "$" + ColumnName(range.Area.StartColumn) + "$" + range.Area.StartRow.ToString(CultureInfo.InvariantCulture);
                if (range.Area.StartRow != range.Area.EndRow || range.Area.StartColumn != range.Area.EndColumn)
                    address += ":$" + ColumnName(range.Area.EndColumn) + "$" + range.Area.EndRow.ToString(CultureInfo.InvariantCulture);
                if (!String.Equals(range.Area.Sheet, _context.Location.Sheet, StringComparison.OrdinalIgnoreCase))
                    address = "'" + range.Area.Sheet.Replace("'", "''") + "'!" + address;
                rendered = address;
                return true;
            }

            private static bool IsSafeUnary(string op)
            {
                return op == "+" || op == "-" || op == "%";
            }

            private static bool IsSafeBinary(string op)
            {
                return op == "+" || op == "-" || op == "*" || op == "/" || op == "^" || op == "&"
                    || op == "=" || op == "<>" || op == "<" || op == ">" || op == "<=" || op == ">=";
            }

            private static bool IsSafeLiteral(string text)
            {
                text = (text ?? String.Empty).Trim();
                if (String.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase)
                    || String.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase)
                    || ErrorLiterals.Contains(text)) return true;
                double number;
                if (Double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)
                    && !Double.IsNaN(number) && !Double.IsInfinity(number)) return true;
                if (text.Length < 2 || text[0] != '"' || text[text.Length - 1] != '"') return false;
                for (int i = 1; i < text.Length - 1; i++)
                {
                    if (text[i] != '"') continue;
                    if (i + 1 >= text.Length - 1 || text[i + 1] != '"') return false;
                    i++;
                }
                return true;
            }
        }

        private static readonly Regex R1C1Cell = new Regex(
            @"^R(?:(?<ra>\d+)|\[(?<rr>[+-]?\d+)\])?C(?:(?<ca>\d+)|\[(?<cr>[+-]?\d+)\])?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static bool TryParseR1C1(string text, ExcelLocation origin, out ReferenceArea area)
        {
            area = null;
            if (String.IsNullOrWhiteSpace(text)) return false;
            string workbook = origin.Workbook;
            string sheet = origin.Sheet;
            string reference = text.Trim();
            int bang = FindQualifierBang(reference);
            if (bang >= 0)
            {
                if (!TryParseQualifier(reference.Substring(0, bang), out workbook, out sheet)) return false;
                workbook = workbook ?? origin.Workbook;
                reference = reference.Substring(bang + 1);
            }
            string[] pieces = reference.Split(':');
            if (pieces.Length < 1 || pieces.Length > 2) return false;
            int originRow, originColumn;
            ReferenceArea originArea;
            if (!ReferenceParser.TryParse(origin.Address, origin.Workbook, origin.Sheet, out originArea)) return false;
            originRow = originArea.StartRow;
            originColumn = originArea.StartColumn;
            int row1, column1, row2, column2;
            if (!TryParseR1C1Cell(pieces[0], originRow, originColumn, out row1, out column1)) return false;
            if (!TryParseR1C1Cell(pieces.Length == 2 ? pieces[1] : pieces[0], originRow, originColumn, out row2, out column2)) return false;
            area = new ReferenceArea(workbook, sheet, Math.Min(row1, row2), Math.Max(row1, row2),
                Math.Min(column1, column2), Math.Max(column1, column2));
            return true;
        }

        private static bool TryParseR1C1Cell(string text, int originRow, int originColumn, out int row, out int column)
        {
            row = column = 0;
            Match match = R1C1Cell.Match(text.Trim());
            if (!match.Success) return false;
            if (!TryR1C1Part(match.Groups["ra"], match.Groups["rr"], originRow, out row)
                || !TryR1C1Part(match.Groups["ca"], match.Groups["cr"], originColumn, out column)) return false;
            return row >= 1 && row <= ReferenceParser.MaxRows && column >= 1 && column <= ReferenceParser.MaxColumns;
        }

        private static bool TryR1C1Part(Group absolute, Group relative, int origin, out int value)
        {
            value = origin;
            if (absolute.Success) return Int32.TryParse(absolute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out value);
            if (relative.Success)
            {
                int offset;
                if (!Int32.TryParse(relative.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out offset)) return false;
                value = origin + offset;
            }
            return true;
        }

        private static int FindQualifierBang(string value)
        {
            bool quoted = false;
            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] == '\'')
                {
                    if (quoted && i + 1 < value.Length && value[i + 1] == '\'') { i++; continue; }
                    quoted = !quoted;
                }
                else if (value[i] == '!' && !quoted) return i;
            }
            return -1;
        }

        private static bool TryParseQualifier(string value, out string workbook, out string sheet)
        {
            workbook = null;
            sheet = null;
            if (value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'')
                value = value.Substring(1, value.Length - 2).Replace("''", "'");
            if (value.Length > 2 && value[0] == '[')
            {
                int close = value.IndexOf(']');
                if (close <= 1 || close == value.Length - 1) return false;
                workbook = value.Substring(1, close - 1);
                sheet = value.Substring(close + 1);
            }
            else sheet = value;
            return !String.IsNullOrEmpty(sheet);
        }

        private sealed class Resolution
        {
            public Resolution(ReferenceArea area, dynamic range) { Area = area; Range = range; }
            public ReferenceArea Area { get; private set; }
            public dynamic Range { get; private set; }
        }

        private sealed class DependencyScan
        {
            private readonly ExcelGateway _gateway;
            private readonly ExcelLocation _source;
            private readonly CancellationToken _cancellation;
            private readonly IProgress<string> _progress;
            private readonly TaskCompletionSource<DependencyResult> _completion;
            private readonly List<DependentCell> _matches = new List<DependentCell>();
            private readonly HashSet<string> _matchKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _warnings = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private dynamic _workbook;
            private ReferenceArea _sourceArea;
            private int _worksheetIndex = 1;
            private int _worksheetCount;
            private dynamic _currentWorksheet;
            private dynamic _formulaCells;
            private int _areaIndex = 1;
            private int _areaCount;
            private dynamic _currentArea;
            private int _areaRows;
            private int _areaColumns;
            private int _nextRow = 1;
            private int _nextColumn = 1;
            private int _formulaCount;
            private bool _finished;

            public DependencyScan(ExcelGateway gateway, ExcelLocation source, CancellationToken cancellation,
                IProgress<string> progress, TaskCompletionSource<DependencyResult> completion)
            {
                _gateway = gateway; _source = source; _cancellation = cancellation; _progress = progress; _completion = completion;
            }

            public void Start()
            {
                _workbook = _gateway.FindOpenWorkbook(_source.Workbook);
                if (!ReferenceParser.TryParse(_source.Address, _source.Workbook, _source.Sheet, out _sourceArea))
                    throw new InvalidOperationException("The selected address is not a supported A1 range.");
                _worksheetCount = Convert.ToInt32(_workbook.Worksheets.Count, CultureInfo.InvariantCulture);
                _progress?.Report("Scanning formulas in " + _worksheetCount.ToString(CultureInfo.InvariantCulture) + " worksheet(s)…");
                ScheduleNext();
            }

            private void ScheduleNext()
            {
                // Input and rendering must run before the next scan batch, especially while
                // ShowDialog is running the WPF dispatcher on Excel's STA.
                _gateway._excelDispatcher.BeginInvoke(DispatcherPriority.Background, (Action)ProcessOneBatch);
            }

            private void ProcessOneBatch()
            {
                try
                {
                    _cancellation.ThrowIfCancellationRequested();
                    if (_finished) return;
                    if (!MoveToNextBlock()) { Finish(); return; }
                    ReadAndProcessCurrentBlock();
                    ScheduleNext();
                }
                catch (OperationCanceledException) { _completion.TrySetCanceled(); }
                catch (Exception ex) { _completion.TrySetException(ex); }
            }

            private bool MoveToNextBlock()
            {
                while (true)
                {
                    if (_currentArea != null && _nextRow <= _areaRows)
                        return true;
                    if (_formulaCells != null && _areaIndex <= _areaCount)
                    {
                        _currentArea = _formulaCells.Areas[_areaIndex++];
                        _areaRows = Convert.ToInt32(_currentArea.Rows.Count, CultureInfo.InvariantCulture);
                        _areaColumns = Convert.ToInt32(_currentArea.Columns.Count, CultureInfo.InvariantCulture);
                        _nextRow = 1;
                        _nextColumn = 1;
                        continue;
                    }
                    if (_worksheetIndex > _worksheetCount) return false;
                    _currentWorksheet = _workbook.Worksheets[_worksheetIndex++];
                    _formulaCells = null; _currentArea = null; _areaIndex = 1; _areaCount = 0;
                    try
                    {
                        _formulaCells = _currentWorksheet.Cells.SpecialCells(XlCellTypeFormulas);
                        _areaCount = Convert.ToInt32(_formulaCells.Areas.Count, CultureInfo.InvariantCulture);
                        _progress?.Report("Scanning " + Convert.ToString(_currentWorksheet.Name, CultureInfo.InvariantCulture) + "…");
                    }
                    catch (COMException ex) when (unchecked((uint)ex.ErrorCode) == 0x800A03EC)
                    {
                        // SpecialCells raises when a worksheet contains no formula cells.
                        _formulaCells = null;
                    }
                }
            }

            private void ReadAndProcessCurrentBlock()
            {
                int rowsLeft = _areaRows - _nextRow + 1;
                int columnsLeft = _areaColumns - _nextColumn + 1;
                int width = Math.Min(columnsLeft, 128);
                int height = Math.Min(rowsLeft, Math.Max(1, MaxCellsPerRead / width));
                dynamic cell = _currentArea.Cells[_nextRow, _nextColumn];
                dynamic block = cell.Resize[height, width];
                object formulas = ReadFormulas(block);
                int row = Convert.ToInt32(cell.Row, CultureInfo.InvariantCulture);
                int column = Convert.ToInt32(cell.Column, CultureInfo.InvariantCulture);
                ProcessFormulaValues(formulas, row, column, height, width);
                _nextColumn += width;
                if (_nextColumn > _areaColumns) { _nextColumn = 1; _nextRow += height; }
            }

            private void ProcessFormulaValues(object formulas, int startRow, int startColumn, int height, int width)
            {
                Array matrix = formulas as Array;
                if (matrix == null)
                {
                    ProcessFormula(formulas as string, startRow, startColumn);
                    return;
                }
                int lowerRow = matrix.GetLowerBound(0);
                int lowerColumn = matrix.GetLowerBound(1);
                for (int r = 0; r < height; r++)
                {
                    for (int c = 0; c < width; c++)
                    {
                        _cancellation.ThrowIfCancellationRequested();
                        object candidate = matrix.GetValue(lowerRow + r, lowerColumn + c);
                        ProcessFormula(candidate as string, startRow + r, startColumn + c);
                    }
                }
            }

            private void ProcessFormula(string formula, int row, int column)
            {
                if (!IsFormula(formula)) return;
                _formulaCount++;
                FormulaParseResult parsed = FormulaParser.Parse(formula);
                if (!parsed.Success || parsed.Root == null)
                {
                    Warn("Some formulas could not be parsed, so the dependency result is incomplete.");
                    return;
                }
                string workbookName = Convert.ToString(_workbook.Name, CultureInfo.InvariantCulture);
                string sheetName = Convert.ToString(_currentWorksheet.Name, CultureInfo.InvariantCulture);
                var location = new ExcelLocation(workbookName, sheetName, ColumnName(column) + row.ToString(CultureInfo.InvariantCulture));
                var context = new InspectionContext(location, formula, String.Empty, true);
                bool matched;
                if (FormulaDependsOn(context, _currentWorksheet, parsed.Root, out matched) && matched)
                {
                    if (_matchKeys.Add(location.Key)) _matches.Add(new DependentCell(context, String.Empty));
                }
            }

            private bool FormulaDependsOn(InspectionContext context, dynamic sheet, FormulaNode node, out bool matched)
            {
                matched = false;
                Resolution resolution;
                string note;
                if (node.Kind == FormulaNodeKind.Reference)
                {
                    if (_gateway.TryResolveReference(context, sheet, node.Text, false, out resolution, out note))
                        matched = _sourceArea.Intersects(resolution.Area);
                    else Warn(note);
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Name)
                {
                    if (_gateway.TryResolveName(context, sheet, node.Text, out resolution, out note))
                        matched = _sourceArea.Intersects(resolution.Area);
                    else Warn(note);
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Binary && node.Operator == ":")
                {
                    if (_gateway.TryResolveRangeOperator(context, sheet, node, out resolution, out note))
                        matched = _sourceArea.Intersects(resolution.Area);
                    else Warn(note);
                    // A resolved range fully replaces its endpoints; do not re-walk a bare
                    // endpoint on the source sheet after a qualified range failed to match.
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Binary && node.Operator == " ")
                {
                    if (TryIntersectionMatches(context, sheet, node, out matched, out note)) return true;
                    Warn(note);
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Unary && node.Operator == "#")
                {
                    if (_gateway.TryResolveSpill(context, sheet, node, out resolution, out note))
                        matched = _sourceArea.Intersects(resolution.Area);
                    else Warn(note);
                    return true;
                }
                if (node.Kind == FormulaNodeKind.Function)
                {
                    string name = (node.Name ?? String.Empty).ToUpperInvariant();
                    if (name == "OFFSET" || name == "INDIRECT")
                    {
                        bool resolved = name == "OFFSET"
                            ? _gateway.TryResolveOffset(context, sheet, node, out resolution, out note)
                            : _gateway.TryResolveIndirect(context, sheet, node, out resolution, out note);
                        if (resolved) matched = _sourceArea.Intersects(resolution.Area);
                        else Warn(note);
                        if (matched) return true;
                    }
                }
                if (node.Children != null)
                {
                    foreach (FormulaNode child in node.Children)
                    {
                        bool childMatch;
                        FormulaDependsOn(context, sheet, child, out childMatch);
                        if (childMatch) { matched = true; return true; }
                    }
                }
                return true;
            }

            private bool TryIntersectionMatches(InspectionContext context, dynamic sheet, FormulaNode node,
                out bool matched, out string note)
            {
                matched = false;
                note = String.Empty;
                if (node.Children == null || node.Children.Count != 2)
                {
                    note = "The reference intersection is incomplete.";
                    return false;
                }
                Resolution left, right;
                if (!_gateway.TryResolveNodeAsRange(context, (object)sheet, node.Children[0], out left, out note)
                    || !_gateway.TryResolveNodeAsRange(context, (object)sheet, node.Children[1], out right, out note)) return false;
                if (!String.Equals(left.Area.Workbook, right.Area.Workbook, StringComparison.OrdinalIgnoreCase)
                    || !String.Equals(left.Area.Sheet, right.Area.Sheet, StringComparison.OrdinalIgnoreCase)) return true;
                int startRow = Math.Max(left.Area.StartRow, right.Area.StartRow);
                int endRow = Math.Min(left.Area.EndRow, right.Area.EndRow);
                int startColumn = Math.Max(left.Area.StartColumn, right.Area.StartColumn);
                int endColumn = Math.Min(left.Area.EndColumn, right.Area.EndColumn);
                if (startRow > endRow || startColumn > endColumn) return true;
                var intersection = new ReferenceArea(left.Area.Workbook, left.Area.Sheet, startRow, endRow, startColumn, endColumn);
                matched = _sourceArea.Intersects(intersection);
                return true;
            }

            private void Warn(string warning)
            {
                if (!String.IsNullOrWhiteSpace(warning)) _warnings.Add(warning);
            }

            private void Finish()
            {
                _finished = true;
                _progress?.Report("Scanned " + _formulaCount.ToString(CultureInfo.InvariantCulture) + " formula cell(s).");
                _completion.TrySetResult(new DependencyResult(_matches.AsReadOnly(), _warnings.ToList().AsReadOnly(), _formulaCount));
            }
        }
    }
}
