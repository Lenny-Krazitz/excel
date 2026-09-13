using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace FormulaNavigator.AddIn.UI
{
    // Constructing WPF in C# avoids the Windows-only markup toolchain during Docker builds.
    internal static class Ui
    {
        internal static readonly Brush Muted = new SolidColorBrush(Color.FromRgb(82, 97, 116));
        internal static readonly Brush Line = new SolidColorBrush(Color.FromRgb(204, 213, 223));
        private static readonly Style FocusBorder = CreateFocusBorder();

        internal static void Window(Window window, string title, double width, double height)
        {
            window.Title = title;
            window.Width = width; window.Height = height;
            window.MinWidth = 300; window.MinHeight = 440;
            window.ShowInTaskbar = false;
            window.ShowActivated = true;
            window.WindowStyle = WindowStyle.ToolWindow;
            // The owner can be either Excel itself or another navigator window.
            // Place once after WPF has created the native HWND, when that relationship
            // and the real (DPI-scaled) window size are both available.
            window.WindowStartupLocation = WindowStartupLocation.Manual;
            window.Loaded += (sender, args) => WindowPlacement.MoveToExcelRightEdge(window);
            window.FontFamily = new FontFamily("Segoe UI"); window.FontSize = 12;
            window.Background = new SolidColorBrush(Color.FromRgb(245, 247, 250));
            KeyboardNavigation.SetTabNavigation(window, KeyboardNavigationMode.Cycle);
            KeyboardNavigation.SetControlTabNavigation(window, KeyboardNavigationMode.Cycle);
        }

        internal static Grid Grid(params GridLength[] rows)
        {
            var grid = new Grid { Margin = new Thickness(8) };
            foreach (var height in rows) grid.RowDefinitions.Add(new RowDefinition { Height = height });
            return grid;
        }

        internal static void Put(Grid grid, UIElement child, int row)
        {
            System.Windows.Controls.Grid.SetRow(child, row);
            grid.Children.Add(child);
        }

        internal static TextBlock Text(string text, bool muted = false)
        {
            var block = new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap, VerticalAlignment = VerticalAlignment.Center };
            if (muted) block.Foreground = Muted;
            return block;
        }

        internal static void LimitText(TextBlock block, double maxHeight)
        {
            block.MaxHeight = maxHeight;
            block.TextTrimming = TextTrimming.CharacterEllipsis;
            block.SetBinding(FrameworkElement.ToolTipProperty,
                new Binding("Text") { RelativeSource = new RelativeSource(RelativeSourceMode.Self) });
        }

        internal static Button Button(string label, RoutedEventHandler click)
        {
            var button = new Button { Content = label, Padding = new Thickness(7, 3, 7, 3), Margin = new Thickness(0, 0, 5, 0), MinHeight = 24 };
            button.Click += click;
            return button;
        }

        internal static StackPanel Horizontal(params UIElement[] children)
        {
            var panel = new StackPanel { Orientation = Orientation.Horizontal };
            foreach (var child in children) panel.Children.Add(child);
            return panel;
        }

        internal static DockPanel Toolbar(UIElement left, UIElement right)
        {
            var panel = new DockPanel();
            DockPanel.SetDock(right, Dock.Right);
            panel.Children.Add(right); panel.Children.Add(left);
            return panel;
        }

        internal static TextBox Formula()
        {
            return new TextBox
            {
                IsReadOnly = true, TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                FontFamily = new FontFamily("Consolas"), FontSize = 12,
                Padding = new Thickness(5), Style = FocusBorder,
                SelectionBrush = new SolidColorBrush(Color.FromRgb(140, 203, 255)),
                SelectionOpacity = 0.8, IsInactiveSelectionHighlightEnabled = true
            };
        }

        internal static TreeView Tree(Type nodeType, string[] properties)
        {
            // A row needs to be readable in a narrow tool window.  Each field has its
            // own wrapping line, so an unusually long calculated value cannot squeeze
            // the node name out of view or require a horizontal scroll bar.
            var panel = new FrameworkElementFactory(typeof(StackPanel));
            panel.SetValue(StackPanel.OrientationProperty, Orientation.Vertical);
            panel.SetValue(FrameworkElement.MarginProperty, new Thickness(1, 2, 1, 2));
            panel.SetBinding(FrameworkElement.ToolTipProperty, new Binding("Note"));
            // The default TreeViewItem template measures its header in an Auto column.
            // Wrapping alone therefore has no effect. Use the arranged container width,
            // which already accounts for indentation, and reserve its expander/padding.
            panel.SetBinding(FrameworkElement.MaxWidthProperty, new Binding("ActualWidth")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(TreeViewItem), 1),
                Converter = new TreeHeaderWidthConverter()
            });

            var label = new FrameworkElementFactory(typeof(TextBlock));
            label.SetValue(TextBlock.FontWeightProperty, FontWeights.SemiBold);
            label.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
            label.SetBinding(TextBlock.TextProperty, new Binding("Label"));
            label.SetBinding(FrameworkElement.ToolTipProperty,
                new Binding(nodeType == typeof(ExplorerNode) ? "Expression.Text" : "Label"));
            panel.AppendChild(label);

            var value = new FrameworkElementFactory(typeof(TextBlock));
            value.SetValue(FrameworkElement.MarginProperty, new Thickness(0, 1, 0, 0));
            value.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
            value.SetValue(TextBlock.TextAlignmentProperty, TextAlignment.Right);
            value.SetBinding(TextBlock.TextProperty, new Binding("Value"));
            value.SetValue(FrameworkElement.StyleProperty, OptionalTreeText());
            value.SetBinding(FrameworkElement.ToolTipProperty, new Binding("Value"));
            panel.AppendChild(value);

            for (int i = 0; i < properties.Length; i++)
            {
                if (properties[i] == "Label" || properties[i] == "Value") continue;
                var text = new FrameworkElementFactory(typeof(TextBlock));
                text.SetValue(FrameworkElement.MarginProperty, new Thickness(0, 1, 0, 0));
                text.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
                text.SetValue(TextBlock.ForegroundProperty, Muted);
                text.SetValue(FrameworkElement.StyleProperty, OptionalTreeText());
                text.SetBinding(TextBlock.TextProperty, new Binding(properties[i]));
                if (properties[i] == "Address" || properties[i] == "Formula" || properties[i] == "Argument")
                    text.SetBinding(FrameworkElement.ToolTipProperty, new Binding(properties[i]));
                panel.AppendChild(text);
            }
            var template = new HierarchicalDataTemplate(nodeType)
            {
                ItemsSource = new Binding("Children"), VisualTree = panel
            };
            var tree = new TreeView { Style = FocusBorder, IsTabStop = true, HorizontalContentAlignment = HorizontalAlignment.Stretch };
            var itemStyle = new Style(typeof(TreeViewItem));
            itemStyle.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch));
            tree.Resources.Add(typeof(TreeViewItem), itemStyle);
            KeyboardNavigation.SetTabNavigation(tree, KeyboardNavigationMode.Once);
            KeyboardNavigation.SetControlTabNavigation(tree, KeyboardNavigationMode.Once);
            KeyboardNavigation.SetDirectionalNavigation(tree, KeyboardNavigationMode.Contained);
            tree.Resources.Add(new DataTemplateKey(nodeType), template);
            VirtualizingStackPanel.SetIsVirtualizing(tree, true);
            VirtualizingStackPanel.SetVirtualizationMode(tree, VirtualizationMode.Recycling);
            ScrollViewer.SetCanContentScroll(tree, true);
            ScrollViewer.SetHorizontalScrollBarVisibility(tree, ScrollBarVisibility.Disabled);
            return tree;
        }

        private static Style OptionalTreeText()
        {
            var style = new Style(typeof(TextBlock));
            style.Setters.Add(new Setter(FrameworkElement.MaxHeightProperty, 34.0));
            style.Setters.Add(new Setter(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis));
            var empty = new Trigger { Property = TextBlock.TextProperty, Value = "" };
            empty.Setters.Add(new Setter(UIElement.VisibilityProperty, Visibility.Collapsed));
            style.Triggers.Add(empty);
            return style;
        }

        private sealed class TreeHeaderWidthConverter : IValueConverter
        {
            public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            {
                double width = value is double ? (double)value : 0;
                return width > 0 ? Math.Max(0, width - 28) : 220;
            }

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            {
                return Binding.DoNothing;
            }
        }

        private static Style CreateFocusBorder()
        {
            var style = new Style(typeof(Control));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, Line));
            var focused = new Trigger { Property = UIElement.IsKeyboardFocusWithinProperty, Value = true };
            focused.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(35, 112, 190))));
            style.Triggers.Add(focused);
            return style;
        }
    }
}
