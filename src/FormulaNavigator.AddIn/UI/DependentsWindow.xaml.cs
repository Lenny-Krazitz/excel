using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using FormulaNavigator.AddIn.Excel;
using FormulaNavigator.Core;

namespace FormulaNavigator.AddIn.UI
{
    public sealed class DependencyRow : INotifyPropertyChanged
    {
        public DependencyRow() { Children = new ObservableCollection<DependencyRow>(); }

        private string _label;
        private string _value;
        private string _formula;
        private string _note;
        private InspectionContext _context;

        public event PropertyChangedEventHandler PropertyChanged;

        private void Changed(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(name));
        }

        public string Label
        {
            get { return _label; }
            set { _label = value; Changed("Label"); }
        }
        public string Value
        {
            get { return _value; }
            set { _value = value; Changed("Value"); }
        }
        public string Formula
        {
            get { return _formula; }
            set { _formula = value; Changed("Formula"); }
        }
        public string Note
        {
            get { return _note; }
            set { _note = value; Changed("Note"); }
        }
        public InspectionContext Context
        {
            get { return _context; }
            set { _context = value; Changed("Context"); }
        }
        public ObservableCollection<DependencyRow> Children { get; private set; }
    }

    public partial class DependentsWindow : Window
    {
        private readonly ExcelGateway gateway;
        private readonly InspectionContext origin;
        private readonly Action<ExcelLocation> explore;
        private readonly List<DependentCell> cells = new List<DependentCell>();
        private readonly TreeNavigation navigation;
        private readonly QueryTiming timing;
        private readonly bool firstQuery;
        private readonly Dictionary<string, DependencyRow> rowsByKey = new Dictionary<string, DependencyRow>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> cellIndexByKey = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private CancellationTokenSource search;
        private Task refreshTask;
        private int refreshGeneration;
        private int refreshCount;
        private bool closePending;
        private bool returnToOrigin = true;
        private ExcelLocation selectedDestination;
        private bool closing;

        public ExcelLocation ExitLocation { get { return returnToOrigin ? origin.Location : selectedDestination; } }

        public DependentsWindow(ExcelGateway gateway, InspectionContext origin, Action<ExcelLocation> explore,
            QueryTiming timing, bool firstQuery)
        {
            this.gateway = gateway;
            this.origin = origin;
            this.explore = explore;
            this.timing = timing;
            this.firstQuery = firstQuery;
            InitializeComponent();
            navigation = new TreeNavigation(this, ResultTree);
            OriginLabel.Text = "Зависят от " + origin.Location;
            FormulaBox.Text = origin.Formula;
        }

        public void CloseWithoutReturning() { returnToOrigin = false; selectedDestination = null; Close(); }

        public void AcceptNavigation(ExcelLocation destination)
        {
            if (closing) return;
            if (destination == null) throw new ArgumentNullException("destination");
            selectedDestination = destination;
            returnToOrigin = false;
            Close();
        }

        private async void Window_Loaded(object sender, RoutedEventArgs args)
        {
            timing.MarkWindowShown();
            await StartRefresh();
        }

        private async void Refresh_Click(object sender, RoutedEventArgs args) { await StartRefresh(); }

        private async Task StartRefresh()
        {
            if (search != null || closing || closePending) return;
            var cancellation = new CancellationTokenSource();
            search = cancellation;
            RefreshButton.IsEnabled = false;
            CancelButton.IsEnabled = true;
            SearchProgress.Visibility = Visibility.Visible;
            StatusLabel.Text = "Ищу зависимые ячейки…";
            WarningsPanel.Visibility = Visibility.Collapsed;
            // The whole refresh, including its finally block, is the task the close
            // logic waits for: when it completes, the search token is already
            // released and the UI is back to the idle state.
            refreshTask = RunRefresh(cancellation);
            await refreshTask;
        }

        private async Task RunRefresh(CancellationTokenSource cancellation)
        {
            try
            {
                await RefreshResults(cancellation.Token);
            }
            finally
            {
                search = null;
                cancellation.Dispose();
                refreshTask = null;
                if (!closing)
                {
                    RefreshButton.IsEnabled = true;
                    CancelButton.IsEnabled = false;
                    SearchProgress.Visibility = Visibility.Collapsed;
                }
            }
        }

        private async Task RefreshResults(CancellationToken cancellation)
        {
            int generation = ++refreshGeneration;
            bool first = refreshCount++ == 0 && firstQuery;
            timing.BeginSearch();
            cells.Clear();
            rowsByKey.Clear();
            cellIndexByKey.Clear();
            ApplyFilter();
            try
            {
                var progress = new Progress<string>(text =>
                {
                    if (!closing && generation == refreshGeneration && !cancellation.IsCancellationRequested)
                        StatusLabel.Text = text;
                });
                DependencyResult result = await gateway.FindDependentsAsync(origin.Location, cancellation, progress);
                if (closing) return;
                cells.AddRange(result.Cells);
                for (int index = 0; index < cells.Count; index++)
                    cellIndexByKey[cells[index].Context.Location.Key] = index;
                timing.RecordTrace(result.TraceMilliseconds);
                timing.RecordSearch(result.ElapsedMilliseconds);
                ApplyFilter();
                // The list is ready here: addresses are rendered and the window is
                // usable while the formulas and values are still loading.
                timing.MarkListReady();
                string searchStage = timing.SearchSeconds();
                string status = "Найдено ячеек: " + cells.Count + ". Штатная трассировка Excel."
                    + (searchStage.Length == 0 ? "" : " Поиск: " + searchStage + " с.")
                    + (first ? " Первый запрос." : " Повторный запрос.");
                if (result.Warnings.Count != 0)
                {
                    WarningsBox.Text = string.Join(Environment.NewLine, result.Warnings);
                    WarningsPanel.Visibility = Visibility.Visible;
                    status = "Результат может быть неполным. " + status;
                }
                StatusLabel.Text = status + " Загружаю формулы и значения…";
                if (result.DetailAreas.Count > 0)
                {
                    await gateway.LoadNativeDetailsAsync(result.DetailAreas, batch => ApplyDetailsAsync(batch, generation),
                        cancellation);
                    if (closing) return;
                    timing.MarkDetailsDone();
                }
                if (generation != refreshGeneration) return;
                string stages = timing.Describe();
                StatusLabel.Text = stages.Length == 0
                    ? status
                    : status + " " + stages;
            }
            catch (OperationCanceledException)
            {
                if (!closing) StatusLabel.Text = "Поиск отменён. Повторный запрос покажет результат заново.";
            }
            catch (Exception error) { if (!closing) StatusLabel.Text = "Поиск не завершён: " + error.Message; }
        }

        // Marshals one detail batch to the UI thread and returns the task that
        // completes when the batch is applied, so the caller's completion means
        // the data is actually visible.
        private Task ApplyDetailsAsync(IReadOnlyList<DependentDetail> batch, int generation)
        {
            var completion = new TaskCompletionSource<bool>();
            Dispatcher.BeginInvoke(new Action(delegate
            {
                try
                {
                    ApplyDetails(batch, generation);
                    completion.SetResult(true);
                }
                catch (Exception error) { completion.SetException(error); }
            }));
            return completion.Task;
        }

        // Details of a superseded request are ignored; every loaded cell updates
        // the data model, and the visible row only when it exists. The rows created
        // by the filter are looked up by key, the model cells by their real index.
        private void ApplyDetails(IReadOnlyList<DependentDetail> batch, int generation)
        {
            if (closing || generation != refreshGeneration) return;
            bool filtered = FilterBox != null && FilterBox.Text.Trim().Length != 0;
            bool rebuild = false;
            foreach (DependentDetail detail in batch)
            {
                string key = detail.Location.Key;
                string formula = detail.Formula;
                var context = new InspectionContext(detail.Location, formula, detail.DisplayValue,
                    formula.StartsWith("=", StringComparison.Ordinal));
                bool unreadable = formula.StartsWith("не прочитано:", StringComparison.Ordinal);
                int index;
                if (cellIndexByKey.TryGetValue(key, out index) && index < cells.Count)
                    cells[index] = new DependentCell(context, unreadable ? formula : String.Empty);
                DependencyRow row;
                if (rowsByKey.TryGetValue(key, out row))
                {
                    if (unreadable) row.Note = formula;
                    row.Formula = formula;
                    row.Value = detail.DisplayValue;
                    row.Context = context;
                    var selected = ResultTree.SelectedItem as DependencyRow;
                    if (selected != null && ReferenceEquals(selected, row)) FormulaBox.Text = formula;
                }
                else if (filtered)
                {
                    // A hidden row may now match the formula filter; rebuild the view
                    // after the whole batch is applied, not in the middle of it.
                    rebuild = true;
                }
            }
            if (rebuild) ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (ResultTree == null) return;
            bool focusResults = IsActive && ResultTree.IsKeyboardFocusWithin;
            string filter = FilterBox == null ? "" : FilterBox.Text.Trim();
            var groups = new List<DependencyRow>();
            rowsByKey.Clear();
            foreach (var group in cells.Where(cell => Matches(cell, filter)).GroupBy(cell => cell.Context.Location.Sheet, StringComparer.OrdinalIgnoreCase))
            {
                var row = new DependencyRow { Label = group.Key + " (" + group.Count() + ")" };
                foreach (var cell in group)
                {
                    var child = new DependencyRow
                    {
                        Label = cell.Context.Location.Address,
                        Value = cell.Context.DisplayValue,
                        Formula = cell.Context.Formula,
                        Note = cell.Note,
                        Context = cell.Context
                    };
                    rowsByKey[cell.Context.Location.Key] = child;
                    row.Children.Add(child);
                }
                groups.Add(row);
            }
            ResultTree.ItemsSource = groups;
            ResultTree.UpdateLayout();
            for (int i = 0; i < groups.Count; i++)
            {
                var item = ResultTree.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                if (item != null) item.IsExpanded = true;
            }
            // A ready selection makes arrows and Enter useful immediately after the search.
            // Rebuilding results while typing must leave keyboard focus in the filter.
            ResultTree.UpdateLayout();
            var firstGroup = groups.Count == 0 ? null : ResultTree.ItemContainerGenerator.ContainerFromIndex(0) as TreeViewItem;
            var firstCell = firstGroup == null ? null : firstGroup.ItemContainerGenerator.ContainerFromIndex(0) as TreeViewItem;
            if (firstCell != null) firstCell.IsSelected = true;
            if (focusResults) navigation.FocusTree();
        }

        private static bool Matches(DependentCell cell, string text)
        {
            if (text.Length == 0) return true;
            return cell.Context.Location.ToString().IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   cell.Context.Formula.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void Filter_Changed(object sender, TextChangedEventArgs args) { ApplyFilter(); }
        private void Cancel_Click(object sender, RoutedEventArgs args) { if (search != null) search.Cancel(); }

        private void Tree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> args)
        {
            navigation.SelectionChanged();
            var row = args.NewValue as DependencyRow;
            if (row == null || row.Context == null) return;
            FormulaBox.Text = row.Context.Formula;
            navigation.Preview(() => gateway.Navigate(row.Context.Location), error => StatusLabel.Text = error.Message);
        }

        private void ExploreSelected()
        {
            navigation.CancelPending();
            var row = ResultTree.SelectedItem as DependencyRow;
            if (row != null && row.Context != null) explore(row.Context.Location);
            if (!closing)
            {
                navigation.FocusTree();
                // Esc from the child inspector resumes the parent's current preview.
                if (row != null && row.Context != null)
                    navigation.Preview(() => gateway.Navigate(row.Context.Location), error => StatusLabel.Text = error.Message);
            }
        }

        private void Explore_Click(object sender, RoutedEventArgs args) { ExploreSelected(); }
        private void Go_Click(object sender, RoutedEventArgs args) { GoToSelected(); }
        private void Close_Click(object sender, RoutedEventArgs args) { Close(); }
        private void Window_KeyDown(object sender, KeyEventArgs args)
        {
            if (args.Key == Key.Escape) { args.Handled = true; Close(); }
            else if (args.Key == Key.Enter && ResultTree.IsKeyboardFocusWithin &&
                (Keyboard.Modifiers == ModifierKeys.None || Keyboard.Modifiers == ModifierKeys.Control))
            {
                args.Handled = true;
                if (Keyboard.Modifiers == ModifierKeys.Control) ExploreSelected();
                else GoToSelected();
            }
        }

        private void GoToSelected()
        {
            var row = ResultTree.SelectedItem as DependencyRow;
            if (row == null || row.Context == null)
            {
                StatusLabel.Text = "Выберите ячейку внутри листа. Ветку можно раскрыть стрелкой вправо.";
                return;
            }
            navigation.CancelPending();
            try
            {
                // Actual navigation runs after the outer modal window releases Excel input.
                gateway.Capture(row.Context.Location);
                AcceptNavigation(row.Context.Location);
            }
            catch (Exception error)
            {
                StatusLabel.Text = error.Message;
                navigation.FocusTree();
            }
        }

        private void Window_Closing(object sender, CancelEventArgs args)
        {
            if (!closePending && refreshTask != null)
            {
                // The window must not close while the search is still restoring the
                // Excel state: the final navigation could otherwise run before the
                // cleanup finishes. Cancel and wait, then close.
                args.Cancel = true;
                closePending = true;
                if (search != null) search.Cancel();
                Task task = refreshTask;
                var dispatcher = Dispatcher;
                Task.Run(delegate
                {
                    try { if (task != null) task.Wait(); } catch (Exception) { }
                    dispatcher.BeginInvoke(new Action(Close));
                });
                return;
            }
            if (closePending && refreshTask != null)
            {
                // A repeated close request must not skip the wait: the deferred
                // close above is the only path allowed to finish the shutdown.
                args.Cancel = true;
                return;
            }
            closing = true;
            navigation.Stop();
            if (search != null) search.Cancel();
            // AddIn applies ExitLocation after the modal message loop has returned.
        }
    }
}
