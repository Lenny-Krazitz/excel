using System;
using System.Windows;
using System.Windows.Interop;
using ExcelDna.Integration;
using FormulaNavigator.AddIn.Excel;
using FormulaNavigator.AddIn.UI;

namespace FormulaNavigator.AddIn
{
    public sealed class AddIn : IExcelAddIn
    {
        private const string MenuName = "Formula Navigator";
        private const string ExploreKey = "^q";
        private const string DependentsKey = "^Q";
        private static ExcelGateway gateway;
        private static object application;
        private static ExplorerWindow explorer;
        private static DependentsWindow dependents;
        private static bool keysRegistered;
        private static bool shuttingDown;
        private static string initializationError;

        public void AutoOpen()
        {
            shuttingDown = false;
            initializationError = null;
            try
            {
                InitializeExcel();
                RegisterKeys();
            }
            catch (Exception error)
            {
                initializationError = error.Message;
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

        [ExcelCommand(Name = "FormulaNavigator_Explore", Description = "Разобрать формулу текущей ячейки",
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

        [ExcelCommand(Name = "FormulaNavigator_Dependents", Description = "Найти зависимые ячейки во всей книге",
            MenuName = MenuName, MenuText = "Зависимые ячейки (Ctrl+Shift+Q)")]
        public static void FindDependents()
        {
            try
            {
                EnsureReady();
                if (ActivateOpenDialog()) return;
                var context = gateway.CaptureActiveCell();
                var window = new DependentsWindow(gateway, context, OpenLocation);
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

        [ExcelCommand(Name = "FormulaNavigator_Status", Description = "Проверить загрузку надстройки",
            MenuName = MenuName, MenuText = "Проверить подключение")]
        public static void ShowStatus()
        {
            try
            {
                string status = "Formula Navigator загружен.\nВерсия: " + typeof(AddIn).Assembly.GetName().Version
                    + "\nФайл: " + ExcelDnaUtil.XllPath
                    + "\n\nДоступ к Excel: " + (gateway != null && !shuttingDown ? "готов" : "не инициализирован")
                    + "\nПоследняя регистрация клавиш: " + (keysRegistered ? "выполнена" : "не выполнена");
                if (!String.IsNullOrEmpty(initializationError)) status += "\n\nОшибка инициализации: " + initializationError;
                if (gateway != null && !shuttingDown) status += "\nИндексация книг: " + gateway.IndexingStatus;
                MessageBox.Show(status, MenuName, MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception error) { Report(error); }
        }

        [ExcelCommand(Name = "FormulaNavigator_RebindKeys", Description = "Повторно назначить Ctrl+Q и Ctrl+Shift+Q",
            MenuName = MenuName, MenuText = "Восстановить горячие клавиши")]
        public static void RebindKeys()
        {
            try
            {
                if (shuttingDown) throw new InvalidOperationException("Надстройка выгружается. Подключите её снова.");
                InitializeExcel();
                RegisterKeys();
                initializationError = null;
                ShowStatus();
            }
            catch (Exception error)
            {
                initializationError = error.Message;
                RestoreKeys();
                Report(error);
            }
        }

        private static void InitializeExcel()
        {
            if (application == null) application = ExcelDnaUtil.Application;
            if (gateway == null) gateway = new ExcelGateway(application);
            gateway.StartIndexing();
        }

        private static void RegisterKeys()
        {
            // Excel's native ON.KEY uses uppercase letters for Ctrl+Shift combinations.
            XlCall.Excel(XlCall.xlcOnKey, ExploreKey, "FormulaNavigator_Explore");
            keysRegistered = true;
            XlCall.Excel(XlCall.xlcOnKey, DependentsKey, "FormulaNavigator_Dependents");
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
                // WPF's modal message loop routes keys before Excel's worksheet message loop.
                window.ShowDialog();
                if (shuttingDown) return;
                if (owner != null)
                {
                    // Enter in the nested inspector completes navigation all the way to Excel.
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
            // Closing a modal window reactivates the previous HWND. Apply the chosen cell
            // afterwards so this does not undo a transition to another workbook/window.
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
            try { XlCall.Excel(XlCall.xlcOnKey, ExploreKey); } catch (Exception) { }
            try { XlCall.Excel(XlCall.xlcOnKey, DependentsKey); } catch (Exception) { }
            keysRegistered = false;
        }

        private static void Report(Exception error)
        {
            if (!shuttingDown)
                MessageBox.Show(error.Message, "Formula Navigator", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
