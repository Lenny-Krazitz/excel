using System.Windows;
using System.Windows.Controls;

namespace FormulaNavigator.AddIn.UI
{
    public partial class ExplorerWindow
    {
        private Button BackButton;
        private TextBlock OriginLabel;
        private TreeView ExpressionTree;
        private TextBox FormulaBox;
        private TextBlock StatusLabel;

        private void InitializeComponent()
        {
            Ui.Window(this, "Formula Navigator — формула", 390, 480);
            PreviewKeyDown += Window_KeyDown;
            Closing += Window_Closing;
            var grid = Ui.Grid(GridLength.Auto, new GridLength(1, GridUnitType.Star),
                new GridLength(5), new GridLength(85), GridLength.Auto, GridLength.Auto);
            grid.RowDefinitions[3].MinHeight = 45;
            Content = grid;

            OriginLabel = Ui.Text("");
            OriginLabel.FontSize = 13; OriginLabel.FontWeight = FontWeights.SemiBold; Ui.LimitText(OriginLabel, 34);
            var subtitle = Ui.Text("Enter — перейти и закрыть · Ctrl+Enter — разобрать ссылку", true);
            subtitle.Margin = new Thickness(0, 2, 5, 0); subtitle.MaxHeight = 34;
            var heading = new StackPanel();
            heading.Children.Add(OriginLabel); heading.Children.Add(subtitle);
            BackButton = Ui.Button("← Назад", Back_Click); BackButton.IsEnabled = false;
            var refresh = Ui.Button("Обновить", Refresh_Click); refresh.Margin = new Thickness(0);
            var toolbar = Ui.Horizontal(BackButton, refresh);
            toolbar.Margin = new Thickness(0, 4, 0, 6);
            heading.Children.Add(toolbar);
            Ui.Put(grid, heading, 0);

            ExpressionTree = Ui.Tree(typeof(ExplorerNode), new[] { "Label", "Argument", "Value", "Address" });
            ExpressionTree.SelectedItemChanged += Tree_SelectedItemChanged;
            ExpressionTree.AddHandler(TreeViewItem.ExpandedEvent, new RoutedEventHandler(Tree_Expanded));
            Ui.Put(grid, ExpressionTree, 1);
            Ui.Put(grid, new GridSplitter { Height = 5, HorizontalAlignment = HorizontalAlignment.Stretch, Focusable = false }, 2);
            FormulaBox = Ui.Formula(); Ui.Put(grid, FormulaBox, 3);
            StatusLabel = Ui.Text("", true); StatusLabel.Margin = new Thickness(0, 5, 0, 5); Ui.LimitText(StatusLabel, 34);
            Ui.Put(grid, StatusLabel, 4);
            var hint = Ui.Text("↑↓ выбор · ←→ ветки · Tab элементы · Alt+← назад · Esc возврат", true);
            hint.Margin = new Thickness(0, 0, 0, 5);
            var close = Ui.Button("Вернуться (Esc)", Close_Click); close.Margin = new Thickness(0);
            var footer = new StackPanel();
            footer.Children.Add(hint);
            var actions = Ui.Horizontal(Ui.Button("Перейти (Enter)", Go_Click), close);
            actions.HorizontalAlignment = HorizontalAlignment.Right;
            footer.Children.Add(actions);
            Ui.Put(grid, footer, 5);
        }
    }
}
