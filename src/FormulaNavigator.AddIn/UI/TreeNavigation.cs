using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace FormulaNavigator.AddIn.UI
{
    // The modal WPF message loop owns keyboard input; keep its focus on the current row.
    internal sealed class TreeNavigation
    {
        private readonly Window window;
        private readonly TreeView tree;
        private TreeViewItem selectedContainer;
        private DispatcherOperation pendingFocus;
        private DispatcherOperation pendingPreview;
        private bool stopped;

        internal TreeNavigation(Window window, TreeView tree)
        {
            this.window = window;
            this.tree = tree;
            tree.AddHandler(TreeViewItem.SelectedEvent, new RoutedEventHandler((sender, args) =>
            {
                selectedContainer = args.OriginalSource as TreeViewItem;
            }));
            tree.GotKeyboardFocus += (sender, args) =>
            {
                // Tab enters the tree at its selection, instead of an arbitrary previous row.
                if (ReferenceEquals(args.OriginalSource, tree)) FocusSelection();
            };
            window.Loaded += Window_Loaded;
            window.PreviewKeyDown += (sender, args) => CancelInitialFocus();
            window.PreviewMouseDown += (sender, args) => CancelInitialFocus();
            window.Closed += (sender, args) => Stop();
        }

        private void Window_Loaded(object sender, RoutedEventArgs args)
        {
            window.Loaded -= Window_Loaded;
            if (stopped) return;
            // Run after all Loaded handlers have created and selected the initial tree item.
            pendingFocus = window.Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() =>
            {
                pendingFocus = null;
                if (stopped || !window.IsVisible || !window.IsActive) return;
                FocusTree();
            }));
        }

        internal void SelectionChanged()
        {
            CancelPreview();
            if (tree.SelectedItem == null) selectedContainer = null;
        }

        internal void Preview(Action navigate, Action<Exception> reportError)
        {
            CancelPreview();
            if (stopped) return;
            var selection = tree.SelectedItem;
            // Let WPF finish selecting/focusing the new row before touching Excel.
            pendingPreview = window.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
            {
                pendingPreview = null;
                if (stopped || !window.IsVisible || !window.IsActive || !tree.IsKeyboardFocusWithin ||
                    !ReferenceEquals(tree.SelectedItem, selection)) return;
                try { navigate(); }
                catch (Exception error) { reportError(error); }
                finally { FocusTree(); }
            }));
        }

        internal void FocusTree()
        {
            if (stopped || !window.IsVisible) return;
            window.Activate();
            if (!window.IsActive) return;
            if (!FocusSelection()) tree.Focus();
        }

        private bool FocusSelection()
        {
            if (stopped || selectedContainer == null || !selectedContainer.IsSelected ||
                !selectedContainer.IsVisible || !ReferenceEquals(selectedContainer.DataContext, tree.SelectedItem)) return false;
            return selectedContainer.Focus();
        }

        internal void CancelPending()
        {
            CancelInitialFocus();
            CancelPreview();
        }

        internal void Stop()
        {
            stopped = true;
            CancelPending();
        }

        private void CancelInitialFocus()
        {
            if (pendingFocus != null) { pendingFocus.Abort(); pendingFocus = null; }
        }

        private void CancelPreview()
        {
            if (pendingPreview != null) { pendingPreview.Abort(); pendingPreview = null; }
        }
    }
}
