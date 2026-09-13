using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace FormulaNavigator.AddIn.UI
{
    public partial class DependentsWindow
    {
        private TextBlock OriginLabel;
        private Button RefreshButton;
        private Button CancelButton;
        private TextBox FilterBox;
        private ProgressBar SearchProgress;
        private TreeView ResultTree;
        private Expander WarningsPanel;
        private TextBox WarningsBox;
        private TextBox FormulaBox;
        private TextBlock StatusLabel;

        private void InitializeComponent()
        {
            Ui.Window(this, "Formula Navigator — зависимые ячейки", 360, 480);
            PreviewKeyDown += Window_KeyDown; Closing += Window_Closing; Loaded += Window_Loaded;
            var grid = Ui.Grid(GridLength.Auto, GridLength.Auto, GridLength.Auto,
                new GridLength(1, GridUnitType.Star), GridLength.Auto, new GridLength(65), GridLength.Auto, GridLength.Auto);
            Content = grid;
            OriginLabel = Ui.Text(""); OriginLabel.FontWeight = FontWeights.SemiBold; OriginLabel.FontSize = 13; Ui.LimitText(OriginLabel, 34);
            var heading = new StackPanel { Margin = new Thickness(0, 0, 0, 6) };
            heading.Children.Add(OriginLabel);
            var subtitle = Ui.Text("Прямые зависимости по всем листам этой книги", true); subtitle.Margin = new Thickness(0, 2, 0, 0); subtitle.MaxHeight = 34;
            heading.Children.Add(subtitle); Ui.Put(grid, heading, 0);
            RefreshButton = Ui.Button("Обновить", Refresh_Click);
            CancelButton = Ui.Button("Отмена", Cancel_Click); CancelButton.IsEnabled = false;
            CancelButton.Margin = new Thickness(0);
            FilterBox = new TextBox { Padding = new Thickness(5, 3, 5, 3), ToolTip = "Фильтр по листу, адресу или формуле" };
            FilterBox.TextChanged += Filter_Changed;
            var buttons = Ui.Horizontal(RefreshButton, CancelButton); buttons.Margin = new Thickness(6, 0, 0, 0);
            var toolbar = Ui.Toolbar(FilterBox, buttons); toolbar.Margin = new Thickness(0, 0, 0, 6);
            Ui.Put(grid, toolbar, 1);
            SearchProgress = new ProgressBar { Height = 3, Margin = new Thickness(0, 0, 0, 5), IsIndeterminate = true, Visibility = Visibility.Collapsed };
            Ui.Put(grid, SearchProgress, 2);
            ResultTree = Ui.Tree(typeof(DependencyRow), new[] { "Label", "Value", "Formula" });
            ResultTree.SelectedItemChanged += Tree_SelectedItemChanged;
            Ui.Put(grid, ResultTree, 3);
            WarningsBox = new TextBox
            {
                IsReadOnly = true, TextWrapping = TextWrapping.Wrap, MaxHeight = 60,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto, BorderThickness = new Thickness(0),
                Padding = new Thickness(5), Background = new SolidColorBrush(Color.FromRgb(255, 244, 220))
            };
            WarningsPanel = new Expander
            {
                Header = "Ограничения поиска", Margin = new Thickness(0, 5, 0, 5),
                Visibility = Visibility.Collapsed, IsExpanded = false, Content = WarningsBox
            };
            Ui.Put(grid, WarningsPanel, 4);
            FormulaBox = Ui.Formula(); FormulaBox.Margin = new Thickness(0, 5, 0, 0);
            Ui.Put(grid, FormulaBox, 5);
            StatusLabel = Ui.Text("", true); StatusLabel.Margin = new Thickness(0, 5, 0, 5); Ui.LimitText(StatusLabel, 34); Ui.Put(grid, StatusLabel, 6);
            var hint = Ui.Text("↑↓ выбор · ←→ ветки · Tab элементы · Enter перейти и закрыть", true);
            hint.Margin = new Thickness(0, 0, 0, 5);
            var close = Ui.Button("Вернуться (Esc)", Close_Click); close.Margin = new Thickness(0);
            var footer = new StackPanel();
            footer.Children.Add(hint);
            var inspect = Ui.Button("Формула (Ctrl+Enter)", Explore_Click);
            var actions = new WrapPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            actions.Children.Add(inspect);
            actions.Children.Add(Ui.Button("Перейти (Enter)", Go_Click));
            actions.Children.Add(close);
            footer.Children.Add(actions);
            Ui.Put(grid, footer, 7);
        }
    }
}
