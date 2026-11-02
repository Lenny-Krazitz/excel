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
        private readonly DependencyResult initialResult;
        private bool returnToOrigin = true;
        private ExcelLocation selectedDestination;
        private bool closing;

        public ExcelLocation ExitLocation { get { return returnToOrigin ? origin.Location : selectedDestination; } }

        // The trace already ran in the command before the window was created, so
        // the window shows the complete list immediately.
        public DependentsWindow(ExcelGateway gateway, InspectionContext origin, Action<ExcelLocation> explore,
            QueryTiming timing, bool firstQuery, DependencyResult result)
        {
            this.gateway = gateway;
            this.origin = origin;
            this.explore = explore;
            this.timing = timing;
            this.firstQuery = firstQuery;
            initialResult = result;
            InitializeComponent();
            navigation = new TreeNavigation(this, ResultTree);
            OriginLabel.Text = "Зависят от " + origin.Location;
            FormulaBox.Text = origin.Formula;
            CancelButton.IsEnabled = false;
            SearchProgress.Visibility = Visibility.Collapsed;
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

        private void Window_Loaded(object sender, RoutedEventArgs args)
        {
            ShowResult(initialResult, firstQuery);
        }

        private void Refresh_Click(object sender, RoutedEventArgs args)
        {
            if (closing) return;
            navigation.CancelPending();
            timing.Restart();
            DependencyResult result;
            try { result = gateway.TraceDependents(origin.Location); }
            catch (Exception error)
            {
                StatusLabel.Text = "Поиск не выполнен: " + error.Message;
                return;
            }
            ShowResult(result, false);
        }

        private void ShowResult(DependencyResult result, bool first)
        {
            cells.Clear();
            if (result != null) cells.AddRange(result.Cells);
            if (result != null) timing.RecordSearch(result.TraceMilliseconds, result.ElapsedMilliseconds);
            ApplyFilter();
            string status = (cells.Count == 0 ? "Зависимых ячеек нет." : "Найдено ячеек: " + cells.Count + ".")
                + " Штатная трассировка Excel." + (first ? " Первый запрос." : " Повторный запрос.");
            if (result != null && result.Warnings.Count != 0)
            {
                WarningsBox.Text = string.Join(Environment.NewLine, result.Warnings);
                WarningsPanel.Visibility = Visibility.Visible;
                WarningsPanel.IsExpanded = true;
                status = "Результат может быть неполным. " + status;
            }
            else WarningsPanel.Visibility = Visibility.Collapsed;
            StatusLabel.Text = status;
            // The total time is taken after the list has been rendered.
            Dispatcher.BeginInvoke(new Action(delegate
            {
                if (closing) return;
                timing.MarkReady();
                string stages = timing.Describe();
                StatusLabel.Text = stages.Length == 0 ? status : status + " " + stages;
            }), System.Windows.Threading.DispatcherPriority.ContextIdle);
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
                {
                    var child = new DependencyRow
                    {
                        Label = cell.Context.Location.Address,
                        Value = cell.Context.DisplayValue,
                        Formula = cell.Context.Formula,
                        Note = cell.Note,
                        Context = cell.Context
                    };
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
        private void Cancel_Click(object sender, RoutedEventArgs args) { }

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
            // AddIn applies ExitLocation after the modal message loop has returned.
        }
    }
}
