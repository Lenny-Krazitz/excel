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

namespace FormulaNavigator.AddIn.UI
{
    public sealed class DependencyRow
    {
        public string Label { get; set; }
        public string Value { get; set; }
        public string Formula { get; set; }
        public string Note { get; set; }
        public InspectionContext Context { get; set; }
        public ObservableCollection<DependencyRow> Children { get; private set; }
        public DependencyRow() { Children = new ObservableCollection<DependencyRow>(); }
    }

    public partial class DependentsWindow : Window
    {
        private const int MaxCachedResults = 16;
        private static readonly Dictionary<string, DependencyResult> ResultCache =
            new Dictionary<string, DependencyResult>(StringComparer.OrdinalIgnoreCase);
        private static readonly Queue<string> ResultCacheOrder = new Queue<string>();

        private readonly ExcelGateway gateway;
        private readonly InspectionContext origin;
        private readonly Action<ExcelLocation> explore;
        private readonly List<DependentCell> cells = new List<DependentCell>();
        private readonly TreeNavigation navigation;
        private CancellationTokenSource search;
        private bool returnToOrigin = true;
        private ExcelLocation selectedDestination;
        private bool closing;

        public ExcelLocation ExitLocation { get { return returnToOrigin ? origin.Location : selectedDestination; } }

        public DependentsWindow(ExcelGateway gateway, InspectionContext origin, Action<ExcelLocation> explore)
        {
            this.gateway = gateway;
            this.origin = origin;
            this.explore = explore;
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
        private async void Window_Loaded(object sender, RoutedEventArgs args) { await RefreshResults(false); }
        private async void Refresh_Click(object sender, RoutedEventArgs args) { await RefreshResults(true); }

        private async Task RefreshResults(bool forceRefresh)
        {
            if (search != null || closing) return;
            var cancellation = new CancellationTokenSource();
            search = cancellation;
            RefreshButton.IsEnabled = false;
            CancelButton.IsEnabled = true;
            SearchProgress.Visibility = Visibility.Visible;
            StatusLabel.Text = "Читаю формулы книги…";
            WarningsPanel.Visibility = Visibility.Collapsed;
            cells.Clear();
            ApplyFilter();
            try
            {
                var progress = new Progress<string>(text => { if (!closing) StatusLabel.Text = text; });
                DependencyResult result = null;
                bool fromCache = !forceRefresh && ResultCache.TryGetValue(origin.Location.Key, out result);
                if (fromCache)
                    progress.Report("Использую сохранённый результат…");
                else
                {
                    result = await gateway.FindDependentsAsync(origin.Location, cancellation.Token, progress);
                    CacheResult(origin.Location.Key, result);
                }
                if (closing) return;
                cells.AddRange(result.Cells);
                ApplyFilter();
                StatusLabel.Text = "Найдено ячеек: " + cells.Count + ". Проверено формул: " + result.FormulaCount
                    + (fromCache ? ". Результат взят из кэша." : ".")
                    + " После изменения книги нажмите «Обновить».";
                if (result.Warnings.Count != 0)
                {
                    WarningsBox.Text = string.Join(Environment.NewLine, result.Warnings);
                    WarningsPanel.Visibility = Visibility.Visible;
                    StatusLabel.Text = "Результат может быть неполным. " + StatusLabel.Text;
                }
                else if (cells.Count == 0)
                    StatusLabel.Text = "Прямые зависимые ячейки в этой книге не найдены. Проверено формул: " + result.FormulaCount + ".";
            }
            catch (OperationCanceledException)
            {
                if (!closing) StatusLabel.Text = "Поиск отменён. Для полного результата нажмите «Обновить».";
            }
            catch (Exception error) { if (!closing) StatusLabel.Text = "Поиск не завершён: " + error.Message; }
            finally
            {
                search = null;
                cancellation.Dispose();
                if (!closing)
                {
                    RefreshButton.IsEnabled = true;
                    CancelButton.IsEnabled = false;
                    SearchProgress.Visibility = Visibility.Collapsed;
                }
            }
        }

        private static void CacheResult(string key, DependencyResult result)
        {
            if (ResultCache.ContainsKey(key))
            {
                ResultCache[key] = result;
                return;
            }
            while (ResultCache.Count >= MaxCachedResults && ResultCacheOrder.Count != 0)
                ResultCache.Remove(ResultCacheOrder.Dequeue());
            ResultCache[key] = result;
            ResultCacheOrder.Enqueue(key);
        }

        private void ApplyFilter()
        {
            if (ResultTree == null) return;
            bool focusResults = IsActive && ResultTree.IsKeyboardFocusWithin;
            string filter = FilterBox == null ? "" : FilterBox.Text.Trim();
            var groups = new List<DependencyRow>();
            foreach (var group in cells.Where(cell => Matches(cell, filter)).GroupBy(cell => cell.Context.Location.Sheet, StringComparer.OrdinalIgnoreCase))
            {
                var row = new DependencyRow { Label = group.Key + " (" + group.Count() + ")" };
                foreach (var cell in group)
                    row.Children.Add(new DependencyRow
                    {
                        Label = cell.Context.Location.Address,
                        Value = cell.Context.DisplayValue,
                        Formula = cell.Context.Formula,
                        Note = cell.Note,
                        Context = cell.Context
                    });
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
            closing = true;
            navigation.Stop();
            if (search != null) search.Cancel();
            // AddIn applies ExitLocation after the modal message loop has returned.
        }
    }
}
