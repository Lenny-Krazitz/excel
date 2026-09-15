using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using FormulaNavigator.Core;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace FormulaNavigator.AddIn.Excel
{
    public sealed partial class ExcelGateway
    {
        // Session state, cursors and all Excel objects belong exclusively to the Excel STA.
        private readonly Dictionary<IntPtr, WorkbookIndex> _indexes = new Dictionary<IntPtr, WorkbookIndex>();
        private ExcelInterop.AppEvents_Event _events;
        private DispatcherTimer _indexTimer;
        private bool _eventsConnected;
        private bool _disposed;
        private bool _pumping;
        private string _eventError;
        private DateTime _reconcileAfter = DateTime.MinValue;
        private IntPtr _priorityWorkbook;

        public string IndexingStatus
        {
            get
            {
                EnsureExcelThread();
                return _eventsConnected
                    ? "готовых индексов: " + _indexes.Values.Count(s => s.Snapshot != null) + "/" + _indexes.Count
                    : "без автоматического обновления: " + _eventError;
            }
        }

        public void StartIndexing()
        {
            EnsureExcelThread();
            if (_indexTimer != null || _disposed) return;
            try
            {
                _events = (ExcelInterop.AppEvents_Event)(object)_application;
                _events.WorkbookOpen += OnWorkbookAvailable;
                _events.NewWorkbook += OnWorkbookAvailable;
                _events.WorkbookActivate += OnWorkbookAvailable;
                _events.SheetActivate += OnSheetActivated;
                _events.SheetChange += OnSheetChanged;
                _events.SheetCalculate += OnSheetCalculated;
                _events.WorkbookNewSheet += OnNewSheet;
                _events.WorkbookBeforeClose += OnBeforeWorkbookClose;
                _eventsConnected = true;
            }
            catch (Exception error)
            {
                _eventError = ComMessage(error);
                DisconnectEvents();
            }
            _indexTimer = new DispatcherTimer(DispatcherPriority.Background, _excelDispatcher);
            _indexTimer.Interval = TimeSpan.FromMilliseconds(50);
            _indexTimer.Tick += PumpIndexes;
            _indexTimer.Start();
        }

        public void Dispose()
        {
            EnsureExcelThread();
            if (_disposed) return;
            _disposed = true;
            if (_indexTimer != null)
            {
                _indexTimer.Stop();
                _indexTimer.Tick -= PumpIndexes;
                _indexTimer = null;
            }
            DisconnectEvents();
            foreach (WorkbookIndex state in _indexes.Values) CloseIndex(state);
            _indexes.Clear();
        }

        private void DisconnectEvents()
        {
            _eventsConnected = false;
            if (_events == null) return;
            // A failed connection may leave only some handlers attached. Remove each independently.
            try { _events.WorkbookOpen -= OnWorkbookAvailable; } catch (Exception) { }
            try { _events.NewWorkbook -= OnWorkbookAvailable; } catch (Exception) { }
            try { _events.WorkbookActivate -= OnWorkbookAvailable; } catch (Exception) { }
            try { _events.SheetActivate -= OnSheetActivated; } catch (Exception) { }
            try { _events.SheetChange -= OnSheetChanged; } catch (Exception) { }
            try { _events.SheetCalculate -= OnSheetCalculated; } catch (Exception) { }
            try { _events.WorkbookNewSheet -= OnNewSheet; } catch (Exception) { }
            try { _events.WorkbookBeforeClose -= OnBeforeWorkbookClose; } catch (Exception) { }
            _events = null;
            // Never ReleaseComObject on the application RCW shared with Excel-DNA.
        }

        private void OnWorkbookAvailable(ExcelInterop.Workbook workbook)
        {
            HandleIndexEvent(() =>
            {
                _priorityWorkbook = ComIdentity(workbook);
                _reconcileAfter = DateTime.MinValue;
            });
        }

        private void OnSheetActivated(object sheet)
        {
            HandleIndexEvent(() =>
            {
                _priorityWorkbook = ComIdentity((object)((dynamic)sheet).Parent);
                _reconcileAfter = DateTime.MinValue;
            });
        }

        private void OnSheetChanged(object sheet, ExcelInterop.Range target)
        {
            HandleIndexEvent(() =>
            {
                WorkbookIndex state;
                if (_indexes.TryGetValue(ComIdentity((object)((dynamic)sheet).Parent), out state))
                    InvalidatePlan(state);
                foreach (WorkbookIndex other in _indexes.Values)
                    if (!ReferenceEquals(other, state)) InvalidateDynamic(other);
            });
        }

        private void OnSheetCalculated(object sheet)
        {
            HandleIndexEvent(() =>
            {
                // A calculation can read other books, volatile functions and external data.
                // Invalidate dynamic addresses across all sessions; static plans remain reusable.
                foreach (WorkbookIndex state in _indexes.Values) InvalidateDynamic(state);
            });
        }

        private void OnNewSheet(ExcelInterop.Workbook workbook, object sheet)
        {
            HandleIndexEvent(() =>
            {
                WorkbookIndex state;
                if (_indexes.TryGetValue(ComIdentity(workbook), out state)) InvalidatePlan(state);
                _reconcileAfter = DateTime.MinValue;
            });
        }

        private void OnBeforeWorkbookClose(ExcelInterop.Workbook workbook, ref bool cancel)
        {
            HandleIndexEvent(() =>
            {
                WorkbookIndex state;
                if (_indexes.TryGetValue(ComIdentity(workbook), out state)) InvalidatePlan(state);
                // Another handler or the save prompt can cancel closing. Check Workbooks later.
                _reconcileAfter = DateTime.UtcNow.AddMilliseconds(500);
            });
        }

        private void HandleIndexEvent(Action action)
        {
            if (_disposed) return;
            try { EnsureExcelThread(); action(); }
            catch (Exception error)
            {
                // Do not let an index failure interrupt an Excel edit. Disable reuse if event
                // processing is unreliable; queries then read a fresh snapshot and show a warning.
                _eventsConnected = false;
                _eventError = ComMessage(error);
                foreach (WorkbookIndex state in _indexes.Values) InvalidatePlan(state);
            }
        }

        private static IntPtr ComIdentity(object value)
        {
            IntPtr identity = Marshal.GetIUnknownForObject(value);
            Marshal.Release(identity);
            return identity;
        }

        private void ReconcileWorkbooks()
        {
            var open = new HashSet<IntPtr>();
            dynamic workbooks = _application.Workbooks;
            int count = Convert.ToInt32(workbooks.Count, CultureInfo.InvariantCulture);
            for (int i = 1; i <= count; i++)
            {
                object workbook = workbooks[i];
                IntPtr identity = ComIdentity(workbook);
                open.Add(identity);
                WorkbookIndex state;
                if (!_indexes.TryGetValue(identity, out state))
                {
                    state = new WorkbookIndex(workbook);
                    _indexes.Add(identity, state);
                }
                string name = Convert.ToString(((dynamic)workbook).Name, CultureInfo.InvariantCulture);
                if (state.Name != name)
                {
                    state.Name = name;
                    InvalidatePlan(state);
                }
            }
            foreach (IntPtr identity in _indexes.Keys.Where(key => !open.Contains(key)).ToArray())
            {
                CloseIndex(_indexes[identity]);
                _indexes.Remove(identity);
            }
            _reconcileAfter = DateTime.UtcNow.AddSeconds(2);
        }

        private static void CloseIndex(WorkbookIndex state)
        {
            state.Closed = true;
            InvalidatePlan(state);
            state.PlanCancellation.Dispose();
            state.DynamicCancellation.Dispose();
            state.Workbook = null;
        }

        private static void InvalidatePlan(WorkbookIndex state)
        {
            state.PlanCancellation.Cancel();
            state.PlanCancellation.Dispose();
            state.PlanCancellation = new CancellationTokenSource();
            state.Reader = null;
            state.Parsing = false;
            state.Snapshot = null;
            state.Error = null;
            state.DirtyAfter = DateTime.UtcNow.AddMilliseconds(750);
            state.Progress = "Ожидаю обновления индекса…";
            InvalidateDynamic(state);
        }

        private static void InvalidateDynamic(WorkbookIndex state)
        {
            if (state.Dynamic == null && state.DynamicReader == null && !state.DynamicBuilding) return;
            if (state.Snapshot != null && state.Snapshot.DynamicFormulas.Count == 0)
            {
                state.Dynamic = EmptyResolvedDependencies;
                return;
            }
            state.DynamicCancellation.Cancel();
            state.DynamicCancellation.Dispose();
            state.DynamicCancellation = new CancellationTokenSource();
            state.DynamicReader = null;
            state.DynamicBuilding = false;
            state.Dynamic = state.Snapshot != null && state.Snapshot.DynamicFormulas.Count == 0
                ? EmptyResolvedDependencies : null;
        }

        private bool CanReadInBackground()
        {
            return _eventsConnected && Convert.ToBoolean(_application.EnableEvents, CultureInfo.InvariantCulture)
                && Convert.ToBoolean(_application.Ready, CultureInfo.InvariantCulture);
        }

        private bool IsCalculationStable()
        {
            int state = Convert.ToInt32(_application.CalculationState, CultureInfo.InvariantCulture);
            // In manual calculation mode use the currently calculated values; never force F9.
            return state == 0 || (state == 2 &&
                Convert.ToInt32(_application.Calculation, CultureInfo.InvariantCulture) == -4135);
        }

        private void PumpIndexes(object sender, EventArgs args)
        {
            if (_disposed || _pumping) return;
            _pumping = true;
            WorkbookIndex selected = null;
            try
            {
                bool demand = _indexes.Values.Any(s => s.Demand != 0);
                if (!demand && !CanReadInBackground()) return;
                if (DateTime.UtcNow >= _reconcileAfter) ReconcileWorkbooks();
                if (!IsCalculationStable()) return;
                var ordered = _indexes.Where(p => !demand || p.Value.Demand != 0)
                    .OrderByDescending(p => p.Value.Demand != 0)
                    .ThenByDescending(p => p.Key == _priorityWorkbook).Select(p => p.Value).ToArray();
                selected = ordered.FirstOrDefault(s => !s.Closed && s.Error == null &&
                    s.Snapshot == null && !s.Parsing && (s.Demand != 0 || DateTime.UtcNow >= s.DirtyAfter));
                if (selected != null)
                {
                    CancellationTokenSource generation = selected.PlanCancellation;
                    if (selected.Reader == null)
                    {
                        string schema = ReadSchema(selected);
                        var reader = new FormulaSnapshotReader(selected.Workbook, selected.Name);
                        if (!ReferenceEquals(selected.PlanCancellation, generation)) return;
                        selected.Schema = schema;
                        selected.Reader = reader;
                    }
                    var clock = Stopwatch.StartNew();
                    do
                    {
                        FormulaSnapshotReader reader = selected.Reader;
                        bool more = reader.ReadNext();
                        if (!ReferenceEquals(selected.PlanCancellation, generation)) return;
                        if (more)
                        {
                            selected.Progress = "Индексирую «" + reader.SheetName + "»: "
                                + reader.Formulas.Count + " формул…";
                        }
                        else
                        {
                            List<InspectionContext> formulas = reader.Formulas;
                            selected.Reader = null;
                            selected.Parsing = true;
                            selected.Progress = "Строю связи: " + formulas.Count + " формул…";
                            BuildSnapshotAsync(selected, formulas, selected.PlanCancellation);
                            break;
                        }
                    } while (clock.ElapsedMilliseconds < 12);
                    return;
                }
                selected = ordered.FirstOrDefault(s => !s.Closed && s.Error == null && s.Snapshot != null
                    && s.Dynamic == null && !s.DynamicBuilding);
                if (selected != null) PumpDynamicIndex(selected);
            }
            catch (Exception error)
            {
                if (selected != null)
                {
                    selected.Reader = null;
                    selected.DynamicReader = null;
                    selected.Error = error;
                }
                // A workbook collection may be temporarily unavailable during open/close.
                _reconcileAfter = DateTime.UtcNow.AddSeconds(2);
            }
            finally { _pumping = false; }
        }

        private static string ReadSchema(WorkbookIndex state)
        {
            dynamic workbook = state.Workbook;
            state.SchemaWarning = null;
            var schema = new StringBuilder();
            AppendSchema(schema, workbook.Name);
            dynamic sheets = workbook.Worksheets;
            int count = Convert.ToInt32(sheets.Count, CultureInfo.InvariantCulture);
            AppendSchema(schema, count);
            for (int i = 1; i <= count; i++)
            {
                dynamic sheet = sheets[i];
                AppendSchema(schema, ComIdentity((object)sheet));
                AppendSchema(schema, sheet.Name);
                try
                {
                    AppendSchema(schema, sheet.UsedRange.Address);
                    dynamic tables = sheet.ListObjects;
                    int tableCount = Convert.ToInt32(tables.Count, CultureInfo.InvariantCulture);
                    AppendSchema(schema, tableCount);
                    for (int t = 1; t <= tableCount; t++)
                    {
                        dynamic table = tables[t];
                        AppendSchema(schema, table.Name);
                        AppendSchema(schema, table.Range.Address);
                        dynamic columns = table.ListColumns;
                        int columnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
                        AppendSchema(schema, columnCount);
                        for (int c = 1; c <= columnCount; c++) AppendSchema(schema, columns[c].Name);
                    }
                }
                catch (Exception)
                {
                    AppendSchema(schema, "tables-unavailable");
                    state.SchemaWarning = "Часть метаданных таблиц недоступна. Каждый поиск перечитывает книгу.";
                }
            }
            // Workbook.Names includes worksheet-local names as well as workbook-scoped names.
            try
            {
                dynamic names = workbook.Names;
                int nameCount = Convert.ToInt32(names.Count, CultureInfo.InvariantCulture);
                AppendSchema(schema, nameCount);
                for (int i = 1; i <= nameCount; i++)
                {
                    try
                    {
                        dynamic name = names.Item(i);
                        AppendSchema(schema, name.Name);
                        AppendSchema(schema, name.RefersTo);
                    }
                    catch (Exception)
                    {
                        AppendSchema(schema, "name-unavailable");
                        state.SchemaWarning = "Часть метаданных имён недоступна. Каждый поиск перечитывает книгу.";
                    }
                }
            }
            catch (Exception)
            {
                AppendSchema(schema, "names-unavailable");
                state.SchemaWarning = "Метаданные имён недоступны. Каждый поиск перечитывает книгу.";
            }
            return schema.ToString();
        }

        private static void AppendSchema(StringBuilder schema, object value)
        {
            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? String.Empty;
            schema.Append(text.Length).Append(':').Append(text);
        }

        private Task<T> OnExcelThread<T>(Func<T> action)
        {
            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            Action run = () =>
            {
                try
                {
                    if (_disposed) throw new ObjectDisposedException(nameof(ExcelGateway));
                    EnsureExcelThread();
                    completion.TrySetResult(action());
                }
                catch (Exception error) { completion.TrySetException(error); }
            };
            if (_excelDispatcher.CheckAccess()) run();
            else
            {
                try
                {
                    DispatcherOperation operation = _excelDispatcher.BeginInvoke(DispatcherPriority.Background, run);
                    operation.Aborted += (sender, args) => completion.TrySetCanceled();
                    if (operation.Status == DispatcherOperationStatus.Aborted) completion.TrySetCanceled();
                }
                catch (Exception error) { completion.TrySetException(error); }
            }
            return completion.Task;
        }

        private sealed class WorkbookIndex
        {
            public WorkbookIndex(object workbook) { Workbook = workbook; }
            public object Workbook;
            public string Name;
            public string Schema;
            public string SchemaWarning;
            public bool Closed;
            public int Demand;
            public DateTime DirtyAfter;
            public string Progress;
            public Exception Error;
            public CancellationTokenSource PlanCancellation = new CancellationTokenSource();
            public CancellationTokenSource DynamicCancellation = new CancellationTokenSource();
            public FormulaSnapshotReader Reader;
            public bool Parsing;
            public DependencySnapshot Snapshot;
            public DynamicSnapshotReader DynamicReader;
            public bool DynamicBuilding;
            public ResolvedDependencies Dynamic;
        }
    }
}
