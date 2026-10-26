using System;
using System.Windows;
using System.Windows.Interop;
using ExcelDna.Integration;
using FormulaNavigator.AddIn.Excel;
using FormulaNavigator.AddIn.UI;
using FormulaNavigator.Core;

namespace FormulaNavigator.AddIn
{
    public sealed class AddIn : IExcelAddIn
    {
        private const string MenuName = "Formula Navigator";
        private static ExcelGateway gateway;
        private static object application;
        private static ExplorerWindow explorer;
        private static DependentsWindow dependents;
        private static bool keysRegistered;
        private static bool shuttingDown;
        private static int dependentQueryCount;

        public void AutoOpen()
        {
            shuttingDown = false;
            try
            {
                InitializeExcel();
                RegisterKeys();
            }
            catch (Exception error)
            {
                RestoreKeys();
                Report(error);
            }
        }

        public void AutoClose()
        {
            shuttingDown = true;
            try
            {
                try { if (explorer != null) explorer.CloseWithoutReturning(); } catch (Exception) { }
                try { if (dependents != null) dependents.CloseWithoutReturning(); } catch (Exception) { }
            }
            finally
            {
                RestoreKeys();
                explorer = null;
                dependents = null;
                if (gateway != null) gateway.Dispose();
                gateway = null;
                application = null;
            }
        }

        [ExcelCommand(Name = AddInCommands.ExploreCommand, Description = "Разобрать формулу текущей ячейки",
            MenuName = MenuName, MenuText = "Формула и зависимости (Ctrl+Q)")]
        public static void Explore()
        {
            try
            {
                EnsureReady();
                if (ActivateOpenDialog()) return;
                ShowExplorer(gateway.CaptureActiveCell());
            }
            catch (Exception error) { Report(error); }
        }

        [ExcelCommand(Name = AddInCommands.DependentsCommand, Description = "Найти прямые зависимые ячейки",
            MenuName = MenuName, MenuText = "Зависимые ячейки (Ctrl+Shift+Q)")]
        public static void FindDependents()
        {
            // The clock starts before any work, including readiness checks, so the
            // measurements cover the full time from the keystroke to the ready window.
            var timing = new QueryTiming();
            try
            {
                EnsureReady();
                if (ActivateOpenDialog()) return;
                var context = gateway.CaptureActiveCell();
                bool first = ++dependentQueryCount == 1;
                var window = new DependentsWindow(gateway, context, OpenLocation, timing, first);
                dependents = window;
                window.Closed += (sender, args) => { if (ReferenceEquals(dependents, window)) dependents = null; };
                try
                {
                    SetExcelOwner(window);
                    window.ShowDialog();
                    NavigateAfterDialog(window.ExitLocation);
                }
                finally
                {
                    if (ReferenceEquals(dependents, window)) dependents = null;
                }
            }
            catch (Exception error) { Report(error); }
        }

        private static void InitializeExcel()
        {
            if (application == null) application = ExcelDnaUtil.Application;
            if (gateway == null) gateway = new ExcelGateway(application);
            gateway.StartIndexing();
        }

        private static void RegisterKeys()
        {
            foreach (CommandKeyBinding binding in AddInCommands.Bindings)
                XlCall.Excel(XlCall.xlcOnKey, binding.Key, binding.Command);
            keysRegistered = true;
        }

        private static void OpenLocation(ExcelLocation location)
        {
            try { ShowExplorer(gateway.Capture(location), dependents); }
            catch (Exception error) { Report(error); }
        }

        private static void ShowExplorer(InspectionContext context, DependentsWindow owner = null)
        {
            if (explorer != null) { explorer.Activate(); return; }
            var window = new ExplorerWindow(gateway, context);
            explorer = window;
            window.Closed += (sender, args) => { if (ReferenceEquals(explorer, window)) explorer = null; };
            try
            {
                if (owner != null) window.Owner = owner;
                else SetExcelOwner(window);
                window.ShowDialog();
                if (shuttingDown) return;
                if (owner != null)
                {
                    if (owner.IsVisible && window.SelectionAccepted) owner.AcceptNavigation(window.ExitLocation);
                }
                else NavigateAfterDialog(window.ExitLocation);
            }
            finally
            {
                if (ReferenceEquals(explorer, window)) explorer = null;
            }
        }

        private static void SetExcelOwner(Window window)
        {
            var handle = gateway.GetWindowHandle();
            if (handle != IntPtr.Zero) new WindowInteropHelper(window).Owner = handle;
        }

        private static bool ActivateOpenDialog()
        {
            Window activeDialog = explorer != null ? (Window)explorer : dependents;
            if (activeDialog == null) return false;
            activeDialog.Activate();
            return true;
        }

        private static void NavigateAfterDialog(ExcelLocation location)
        {
            if (shuttingDown || location == null) return;
            gateway.Navigate(location);
        }

        private static void EnsureReady()
        {
            if (shuttingDown || gateway == null)
                throw new InvalidOperationException("Надстройка не загружена. Подключите FormulaNavigator64.xll в параметрах Excel.");
        }

        private static void RestoreKeys()
        {
            if (!keysRegistered) return;
            foreach (CommandKeyBinding binding in AddInCommands.Bindings)
            {
                try { XlCall.Excel(XlCall.xlcOnKey, binding.Key); } catch (Exception) { }
            }
            keysRegistered = false;
        }

        private static void Report(Exception error)
        {
            if (!shuttingDown)
                MessageBox.Show(error.Message, "Formula Navigator", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
