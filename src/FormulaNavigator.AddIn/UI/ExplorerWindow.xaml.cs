using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using FormulaNavigator.AddIn.Excel;
using FormulaNavigator.Core;

namespace FormulaNavigator.AddIn.UI
{
    public partial class ExplorerWindow : Window
    {
        private readonly ExcelGateway gateway;
        private readonly ExcelLocation origin;
        private readonly Stack<InspectionContext> history = new Stack<InspectionContext>();
        private readonly TreeNavigation navigation;
        private InspectionContext current;
        private bool returnToOrigin = true;
        private ExcelLocation selectedDestination;
        private string parseStatus = "";

        public ExcelLocation ExitLocation { get { return returnToOrigin ? origin : selectedDestination; } }
        public bool SelectionAccepted { get { return selectedDestination != null; } }

        public ExplorerWindow(ExcelGateway gateway, InspectionContext context)
        {
            this.gateway = gateway;
            origin = context.Location;
            InitializeComponent();
            navigation = new TreeNavigation(this, ExpressionTree);
            SetContext(context);
            Loaded += (sender, args) => SelectRoot(false);
        }

        public void CloseWithoutReturning()
        {
            returnToOrigin = false;
            selectedDestination = null;
            Close();
        }

        private void SetContext(InspectionContext context)
        {
            current = context;
            OriginLabel.Text = context.Location.ToString();
            FormulaBox.Text = context.HasFormula ? context.Formula : "(В этой ячейке нет формулы.)";
            FormulaNode expression = null;
            bool parsedSuccessfully = false;
            parseStatus = "Значения показаны на момент просмотра. Обновить — перечитать ячейку.";
            if (context.HasFormula)
            {
                var parsed = FormulaParser.Parse(context.Formula);
                expression = parsed.Root;
                parsedSuccessfully = parsed.Success;
                if (!parsed.Success)
                    parseStatus = "Разбор неполный: " + string.Join("; ", parsed.Diagnostics);
            }
            var root = ExplorerNode.CreateRoot(context, expression);
            if (parsedSuccessfully && root.Children.Count == 0)
                parseStatus = "В этой формуле нет ссылок на ячейки. Исходный текст показан ниже.";
            ExpressionTree.ItemsSource = new[] { root };
            StatusLabel.Text = parseStatus;
            BackButton.IsEnabled = history.Count != 0;
        }

        private void SelectRoot(bool focus = true)
        {
            ExpressionTree.UpdateLayout();
            var item = ExpressionTree.ItemContainerGenerator.ContainerFromIndex(0) as TreeViewItem;
            if (item == null) return;
            item.IsExpanded = true;
            item.IsSelected = true;
            if (focus) navigation.FocusTree();
        }

        private void Tree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> args)
        {
            navigation.SelectionChanged();
            var node = args.NewValue as ExplorerNode;
            if (node == null) return;
            try
            {
                node.Inspect(gateway);
                FormulaBox.Text = node.Context.HasFormula ? node.Context.Formula : "(В этой ячейке нет формулы.)";
                if (node.Expression != null)
                {
                    int start = Math.Max(0, Math.Min(node.Expression.Start, FormulaBox.Text.Length));
                    int length = Math.Max(0, Math.Min(node.Expression.Length, FormulaBox.Text.Length - start));
                    FormulaBox.Select(start, length);
                    int line = FormulaBox.GetLineIndexFromCharacterIndex(start);
                    if (line >= 0) FormulaBox.ScrollToLine(line);
                }
                else FormulaBox.Select(0, 0);
                StatusLabel.Text = string.IsNullOrEmpty(node.Note) ? parseStatus : node.Note;
                if (node.Target != null)
                    navigation.Preview(() => gateway.Navigate(node.Target), error => StatusLabel.Text = error.Message);
            }
            catch (Exception error) { StatusLabel.Text = error.Message; }
        }

        private void Tree_Expanded(object sender, RoutedEventArgs args)
        {
            var container = args.OriginalSource as TreeViewItem;
            var node = container == null ? null : container.DataContext as ExplorerNode;
            if (node == null) return;
            try
            {
                // Evaluation is lazy: opening a formula never evaluates its entire tree.
                node.Inspect(gateway);
                int count = 0;
                foreach (var child in node.Children)
                {
                    if (++count > 64) break;
                    child.Inspect(gateway);
                }
            }
            catch (Exception error) { StatusLabel.Text = error.Message; }
        }

        private void DrillIntoSelected()
        {
            var node = ExpressionTree.SelectedItem as ExplorerNode;
            if (node == null) return;
            try
            {
                node.Inspect(gateway);
                if (node.Target == null)
                {
                    StatusLabel.Text = "У этого узла нет адреса для перехода. Ветку можно раскрыть стрелкой вправо.";
                    return;
                }
                var next = gateway.Capture(node.Target);
                if (string.Equals(next.Location.Key, current.Location.Key, StringComparison.OrdinalIgnoreCase)) return;
                if (history.Count >= 128)
                {
                    StatusLabel.Text = "Достигнут предел истории (128 переходов). Вернитесь назад или откройте окно заново.";
                    return;
                }
                history.Push(current);
                SetContext(next);
                SelectRoot();
            }
            catch (Exception error) { StatusLabel.Text = error.Message; }
        }

        private void Window_KeyDown(object sender, KeyEventArgs args)
        {
            if (args.Key == Key.Escape) { args.Handled = true; Close(); }
            else if (args.Key == Key.System && args.SystemKey == Key.Left && (Keyboard.Modifiers & ModifierKeys.Alt) != 0)
            { args.Handled = true; Back(); }
            else if (args.Key == Key.Enter && ExpressionTree.IsKeyboardFocusWithin &&
                (Keyboard.Modifiers == ModifierKeys.None || Keyboard.Modifiers == ModifierKeys.Control))
            {
                args.Handled = true;
                if (Keyboard.Modifiers == ModifierKeys.Control) DrillIntoSelected();
                else GoToSelected();
            }
        }

        private void GoToSelected()
        {
            var node = ExpressionTree.SelectedItem as ExplorerNode;
            if (node == null) return;
            navigation.CancelPending();
            try
            {
                node.Inspect(gateway);
                if (node.Target == null)
                {
                    StatusLabel.Text = "У этого узла нет адреса. Выберите ссылку на ячейку или раскройте ветку стрелкой вправо.";
                    return;
                }
                // Check that the address still exists without activating disabled Excel.
                gateway.Capture(node.Target);
                selectedDestination = node.Target;
                returnToOrigin = false;
                Close();
            }
            catch (Exception error)
            {
                StatusLabel.Text = error.Message;
                navigation.FocusTree();
            }
        }

        private void Back()
        {
            if (history.Count == 0) return;
            SetContext(history.Pop());
            SelectRoot();
        }

        private void Back_Click(object sender, RoutedEventArgs args) { Back(); }
        private void Refresh_Click(object sender, RoutedEventArgs args)
        {
            try { SetContext(gateway.Capture(current.Location)); SelectRoot(); }
            catch (Exception error) { StatusLabel.Text = error.Message; }
        }
        private void Go_Click(object sender, RoutedEventArgs args) { GoToSelected(); }
        private void Close_Click(object sender, RoutedEventArgs args) { Close(); }
        private void Window_Closing(object sender, CancelEventArgs args)
        {
            navigation.Stop();
            // AddIn applies ExitLocation after ShowDialog has re-enabled the owner windows.
        }
    }
}
