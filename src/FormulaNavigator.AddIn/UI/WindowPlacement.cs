using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace FormulaNavigator.AddIn.UI
{
    internal static class WindowPlacement
    {
        private const uint MonitorDefaultToNearest = 2;
        private const uint SwpNoSize = 0x0001;
        private const uint SwpNoZOrder = 0x0004;
        private const uint SwpNoActivate = 0x0010;

        internal static void MoveToExcelRightEdge(Window window)
        {
            try
            {
                var windowHandle = new WindowInteropHelper(window).Handle;
                var excelHandle = FindExcelOwner(window);
                if (windowHandle == IntPtr.Zero || excelHandle == IntPtr.Zero) return;

                var monitor = MonitorFromWindow(excelHandle, MonitorDefaultToNearest);
                if (monitor == IntPtr.Zero) return;
                var monitorInfo = new MonitorInfo { cbSize = Marshal.SizeOf(typeof(MonitorInfo)) };
                if (!GetMonitorInfo(monitor, ref monitorInfo)) return;

                Rect ownerBounds;
                if (!GetWindowRect(excelHandle, out ownerBounds)) return;
                // Align to the right edge of Excel, while keeping the entire tool
                // window in the monitor work area (including negative co-ordinates).
                int desiredRight = Math.Max(monitorInfo.rcWork.Left,
                    Math.Min(ownerBounds.Right, monitorInfo.rcWork.Right));
                for (int pass = 0; pass < 2; pass++)
                {
                    Rect windowBounds;
                    if (!GetWindowRect(windowHandle, out windowBounds)) return;
                    int width = windowBounds.Right - windowBounds.Left;
                    int height = windowBounds.Bottom - windowBounds.Top;
                    int x = Math.Max(monitorInfo.rcWork.Left,
                        Math.Min(desiredRight - width, monitorInfo.rcWork.Right - width));
                    int y = Math.Max(monitorInfo.rcWork.Top, Math.Min(ownerBounds.Top, monitorInfo.rcWork.Bottom - height));
                    if (windowBounds.Left == x && windowBounds.Top == y) break;
                    SetWindowPos(windowHandle, IntPtr.Zero, x, y, 0, 0, SwpNoSize | SwpNoZOrder | SwpNoActivate);
                    // Moving across monitors can resize WPF through WM_DPICHANGED.
                    // Re-read the native size before aligning its right edge again.
                }
            }
            catch (Exception)
            {
                // Placement is a convenience only; opening the navigator must still work
                // on unusual window managers or remote-desktop configurations.
            }
        }

        private static IntPtr FindExcelOwner(Window window)
        {
            Window root = window;
            while (root.Owner != null) root = root.Owner;
            return new WindowInteropHelper(root).Owner;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr handle, uint flags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo monitorInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr handle, out Rect rectangle);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr handle, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            internal int Left;
            internal int Top;
            internal int Right;
            internal int Bottom;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct MonitorInfo
        {
            internal int cbSize;
            internal Rect rcMonitor;
            internal Rect rcWork;
            internal uint dwFlags;
        }
    }
}
