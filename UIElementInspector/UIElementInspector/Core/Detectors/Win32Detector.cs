using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Detectors
{
    /// <summary>
    /// Detects elements using Win32 API for native Windows applications (SPY++ functionality)
    /// </summary>
    public class Win32Detector : IElementDetector
    {
        #region Win32 API Declarations

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsZoomed(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowUnicode(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetMenu(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetDlgCtrlID(IntPtr hwndCtl);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        // GetWindowLong indexes
        private const int GWL_STYLE = -16;
        private const int GWL_EXSTYLE = -20;
        private const int GWL_WNDPROC = -4;
        private const int GWL_HINSTANCE = -6;
        private const int GWL_ID = -12;

        // GetWindow commands
        private const uint GW_CHILD = 5;
        private const uint GW_HWNDNEXT = 2;
        private const uint GW_HWNDPREV = 3;

        // Window Styles (WS_*)
        private const uint WS_OVERLAPPED = 0x00000000;
        private const uint WS_POPUP = 0x80000000;
        private const uint WS_CHILD = 0x40000000;
        private const uint WS_MINIMIZE = 0x20000000;
        private const uint WS_VISIBLE = 0x10000000;
        private const uint WS_DISABLED = 0x08000000;
        private const uint WS_CLIPSIBLINGS = 0x04000000;
        private const uint WS_CLIPCHILDREN = 0x02000000;
        private const uint WS_MAXIMIZE = 0x01000000;
        private const uint WS_CAPTION = 0x00C00000;
        private const uint WS_BORDER = 0x00800000;
        private const uint WS_DLGFRAME = 0x00400000;
        private const uint WS_VSCROLL = 0x00200000;
        private const uint WS_HSCROLL = 0x00100000;
        private const uint WS_SYSMENU = 0x00080000;
        private const uint WS_THICKFRAME = 0x00040000;
        private const uint WS_GROUP = 0x00020000;
        private const uint WS_TABSTOP = 0x00010000;

        // Extended Window Styles (WS_EX_*)
        private const uint WS_EX_DLGMODALFRAME = 0x00000001;
        private const uint WS_EX_NOPARENTNOTIFY = 0x00000004;
        private const uint WS_EX_TOPMOST = 0x00000008;
        private const uint WS_EX_ACCEPTFILES = 0x00000010;
        private const uint WS_EX_TRANSPARENT = 0x00000020;
        private const uint WS_EX_MDICHILD = 0x00000040;
        private const uint WS_EX_TOOLWINDOW = 0x00000080;
        private const uint WS_EX_WINDOWEDGE = 0x00000100;
        private const uint WS_EX_CLIENTEDGE = 0x00000200;
        private const uint WS_EX_CONTEXTHELP = 0x00000400;
        private const uint WS_EX_RIGHT = 0x00001000;
        private const uint WS_EX_RTLREADING = 0x00002000;
        private const uint WS_EX_LEFTSCROLLBAR = 0x00004000;
        private const uint WS_EX_CONTROLPARENT = 0x00010000;
        private const uint WS_EX_STATICEDGE = 0x00020000;
        private const uint WS_EX_APPWINDOW = 0x00040000;
        private const uint WS_EX_LAYERED = 0x00080000;

        #endregion

        public string Name => "Win32 API";

        public bool CanDetect(System.Windows.Point screenPoint)
        {
            try
            {
                var point = new POINT { X = (int)screenPoint.X, Y = (int)screenPoint.Y };
                var hwnd = WindowFromPoint(point);
                return hwnd != IntPtr.Zero;
            }
            catch
            {
                return false;
            }
        }

        public async Task<ElementInfo> GetElementAtPoint(System.Windows.Point screenPoint, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var stopwatch = Stopwatch.StartNew();
                    var point = new POINT { X = (int)screenPoint.X, Y = (int)screenPoint.Y };
                    var hwnd = WindowFromPoint(point);

                    if (hwnd == IntPtr.Zero)
                        return null;

                    var info = ExtractElementInfo(hwnd, profile);
                    stopwatch.Stop();
                    info.CollectionDuration = stopwatch.Elapsed;

                    return info;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Win32 detection error: {ex.Message}");
                    return null;
                }
            });
        }

        public async Task<List<ElementInfo>> GetAllElements(IntPtr windowHandle, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                var elements = new List<ElementInfo>();
                try
                {
                    if (windowHandle == IntPtr.Zero)
                    {
                        // Get all top-level windows - not practical, skip
                        return elements;
                    }

                    // Get element for main window
                    var mainInfo = ExtractElementInfo(windowHandle, profile);
                    elements.Add(mainInfo);

                    // Enumerate all child windows
                    var childHandles = new List<IntPtr>();
                    EnumChildWindows(windowHandle, (hWnd, lParam) =>
                    {
                        childHandles.Add(hWnd);
                        return true;
                    }, IntPtr.Zero);

                    foreach (var childHwnd in childHandles)
                    {
                        try
                        {
                            var childInfo = ExtractElementInfo(childHwnd, profile);
                            elements.Add(childInfo);
                        }
                        catch { }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Win32 GetAllElements error: {ex.Message}");
                }
                return elements;
            });
        }

        public async Task<List<ElementInfo>> GetElementsInRegion(Rect region, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                var elements = new List<ElementInfo>();
                try
                {
                    // Check center point
                    var centerPoint = new POINT
                    {
                        X = (int)(region.X + region.Width / 2),
                        Y = (int)(region.Y + region.Height / 2)
                    };

                    var hwnd = WindowFromPoint(centerPoint);
                    if (hwnd != IntPtr.Zero)
                    {
                        var info = ExtractElementInfo(hwnd, profile);

                        // Check if element is in region
                        if (info.BoundingRectangle.IntersectsWith(region))
                        {
                            elements.Add(info);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Win32 GetElementsInRegion error: {ex.Message}");
                }
                return elements;
            });
        }

        public Task<ElementInfo> GetElementTree(ElementInfo rootElement, CollectionProfile profile)
        {
            return Task.FromResult(rootElement);
        }

        public Task<ElementInfo> RefreshElement(ElementInfo element, CollectionProfile profile)
        {
            return Task.Run(() =>
            {
                try
                {
                    if (element.Win32_HWND != IntPtr.Zero)
                    {
                        return ExtractElementInfo(element.Win32_HWND, profile);
                    }
                    else if (element.NativeWindowHandle != IntPtr.Zero)
                    {
                        return ExtractElementInfo(element.NativeWindowHandle, profile);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Win32 RefreshElement error: {ex.Message}");
                }
                return element;
            });
        }

        private ElementInfo ExtractElementInfo(IntPtr hwnd, CollectionProfile profile)
        {
            var info = new ElementInfo
            {
                DetectionMethod = Name,
                CollectionProfile = profile.ToString(),
                CaptureTime = DateTime.Now
            };

            try
            {
                // Basic handle info
                info.Win32_HWND = hwnd;
                info.NativeWindowHandle = hwnd;
                info.WindowHandle = hwnd;

                // Parent handle
                info.Win32_ParentHWND = GetParent(hwnd);

                // Window class name
                var className = new StringBuilder(256);
                GetClassName(hwnd, className, className.Capacity);
                info.Win32_ClassName = className.ToString();
                info.ClassName = info.Win32_ClassName;
                info.WindowClassName = info.Win32_ClassName;

                // Window text (caption)
                int textLength = GetWindowTextLength(hwnd);
                if (textLength > 0)
                {
                    var windowText = new StringBuilder(textLength + 1);
                    GetWindowText(hwnd, windowText, windowText.Capacity);
                    info.Win32_WindowText = windowText.ToString();
                    info.Name = info.Win32_WindowText;
                    info.WindowTitle = info.Win32_WindowText;
                }

                // Process and Thread IDs
                uint processId;
                info.Win32_ThreadId = (int)GetWindowThreadProcessId(hwnd, out processId);
                info.ProcessId = (int)processId;

                // Window and Client rectangles
                RECT windowRect, clientRect;
                if (GetWindowRect(hwnd, out windowRect))
                {
                    info.Win32_WindowRect = new Rect(
                        windowRect.Left, windowRect.Top,
                        windowRect.Right - windowRect.Left,
                        windowRect.Bottom - windowRect.Top);
                    info.BoundingRectangle = info.Win32_WindowRect;
                    info.X = windowRect.Left;
                    info.Y = windowRect.Top;
                    info.Width = windowRect.Right - windowRect.Left;
                    info.Height = windowRect.Bottom - windowRect.Top;
                }

                if (GetClientRect(hwnd, out clientRect))
                {
                    info.Win32_ClientRect = new Rect(
                        0, 0,
                        clientRect.Right - clientRect.Left,
                        clientRect.Bottom - clientRect.Top);
                    info.ClientRect = info.Win32_ClientRect;
                }

                // Window states
                info.Win32_IsVisible = IsWindowVisible(hwnd);
                info.Win32_IsEnabled = IsWindowEnabled(hwnd);
                info.Win32_IsMaximized = IsZoomed(hwnd);
                info.Win32_IsMinimized = IsIconic(hwnd);
                info.Win32_IsUnicode = IsWindowUnicode(hwnd);
                info.IsVisible = info.Win32_IsVisible ?? false;
                info.IsEnabled = info.Win32_IsEnabled ?? true;

                // Window styles
                info.Win32_WindowStyles = (uint)GetWindowLong(hwnd, GWL_STYLE);
                info.Win32_WindowStyles_Parsed = ParseWindowStyles(info.Win32_WindowStyles);

                // Extended styles
                info.Win32_ExtendedStyles = (uint)GetWindowLong(hwnd, GWL_EXSTYLE);
                info.Win32_ExtendedStyles_Parsed = ParseExtendedStyles(info.Win32_ExtendedStyles);

                // Menu handle
                info.Win32_MenuHandle = GetMenu(hwnd);

                // Instance handle
                info.Win32_InstanceHandle = GetWindowLongPtr(hwnd, GWL_HINSTANCE);

                // WndProc (window procedure)
                info.Win32_WndProc = GetWindowLongPtr(hwnd, GWL_WNDPROC);

                // Control ID
                int ctrlId = GetDlgCtrlID(hwnd);
                if (ctrlId != 0)
                {
                    info.Win32_ControlId = ctrlId.ToString();
                }

                // Get child windows count
                if (profile == CollectionProfile.Full)
                {
                    var childCount = 0;
                    EnumChildWindows(hwnd, (hWnd, lParam) =>
                    {
                        if (GetParent(hWnd) == hwnd) // Direct children only
                            childCount++;
                        return true;
                    }, IntPtr.Zero);
                    info.CustomProperties["Win32_DirectChildCount"] = childCount;

                    // Get all child HWNDs
                    EnumChildWindows(hwnd, (hWnd, lParam) =>
                    {
                        if (GetParent(hWnd) == hwnd)
                            info.Win32_ChildHWNDs.Add(hWnd);
                        return true;
                    }, IntPtr.Zero);
                }

                // Set element type based on class name
                info.ElementType = DetermineElementType(info.Win32_ClassName);

                // Process info
                try
                {
                    var process = Process.GetProcessById(info.ProcessId);
                    info.ApplicationName = process.ProcessName;
                    try
                    {
                        info.ApplicationPath = process.MainModule?.FileName;
                    }
                    catch { }
                }
                catch { }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Win32 extraction error: {ex.Message}");
            }

            return info;
        }

        private IntPtr GetWindowLongPtr(IntPtr hwnd, int nIndex)
        {
            if (IntPtr.Size == 8)
                return GetWindowLongPtr64(hwnd, nIndex);
            else
                return new IntPtr(GetWindowLong(hwnd, nIndex));
        }

        private string ParseWindowStyles(uint styles)
        {
            var result = new List<string>();

            if ((styles & WS_POPUP) != 0) result.Add("WS_POPUP");
            if ((styles & WS_CHILD) != 0) result.Add("WS_CHILD");
            if ((styles & WS_MINIMIZE) != 0) result.Add("WS_MINIMIZE");
            if ((styles & WS_VISIBLE) != 0) result.Add("WS_VISIBLE");
            if ((styles & WS_DISABLED) != 0) result.Add("WS_DISABLED");
            if ((styles & WS_CLIPSIBLINGS) != 0) result.Add("WS_CLIPSIBLINGS");
            if ((styles & WS_CLIPCHILDREN) != 0) result.Add("WS_CLIPCHILDREN");
            if ((styles & WS_MAXIMIZE) != 0) result.Add("WS_MAXIMIZE");
            if ((styles & WS_CAPTION) == WS_CAPTION) result.Add("WS_CAPTION");
            if ((styles & WS_BORDER) != 0) result.Add("WS_BORDER");
            if ((styles & WS_DLGFRAME) != 0) result.Add("WS_DLGFRAME");
            if ((styles & WS_VSCROLL) != 0) result.Add("WS_VSCROLL");
            if ((styles & WS_HSCROLL) != 0) result.Add("WS_HSCROLL");
            if ((styles & WS_SYSMENU) != 0) result.Add("WS_SYSMENU");
            if ((styles & WS_THICKFRAME) != 0) result.Add("WS_THICKFRAME");
            if ((styles & WS_GROUP) != 0) result.Add("WS_GROUP");
            if ((styles & WS_TABSTOP) != 0) result.Add("WS_TABSTOP");

            return result.Count > 0 ? string.Join(" | ", result) : "WS_OVERLAPPED";
        }

        private string ParseExtendedStyles(uint styles)
        {
            var result = new List<string>();

            if ((styles & WS_EX_DLGMODALFRAME) != 0) result.Add("WS_EX_DLGMODALFRAME");
            if ((styles & WS_EX_NOPARENTNOTIFY) != 0) result.Add("WS_EX_NOPARENTNOTIFY");
            if ((styles & WS_EX_TOPMOST) != 0) result.Add("WS_EX_TOPMOST");
            if ((styles & WS_EX_ACCEPTFILES) != 0) result.Add("WS_EX_ACCEPTFILES");
            if ((styles & WS_EX_TRANSPARENT) != 0) result.Add("WS_EX_TRANSPARENT");
            if ((styles & WS_EX_MDICHILD) != 0) result.Add("WS_EX_MDICHILD");
            if ((styles & WS_EX_TOOLWINDOW) != 0) result.Add("WS_EX_TOOLWINDOW");
            if ((styles & WS_EX_WINDOWEDGE) != 0) result.Add("WS_EX_WINDOWEDGE");
            if ((styles & WS_EX_CLIENTEDGE) != 0) result.Add("WS_EX_CLIENTEDGE");
            if ((styles & WS_EX_CONTEXTHELP) != 0) result.Add("WS_EX_CONTEXTHELP");
            if ((styles & WS_EX_RIGHT) != 0) result.Add("WS_EX_RIGHT");
            if ((styles & WS_EX_RTLREADING) != 0) result.Add("WS_EX_RTLREADING");
            if ((styles & WS_EX_LEFTSCROLLBAR) != 0) result.Add("WS_EX_LEFTSCROLLBAR");
            if ((styles & WS_EX_CONTROLPARENT) != 0) result.Add("WS_EX_CONTROLPARENT");
            if ((styles & WS_EX_STATICEDGE) != 0) result.Add("WS_EX_STATICEDGE");
            if ((styles & WS_EX_APPWINDOW) != 0) result.Add("WS_EX_APPWINDOW");
            if ((styles & WS_EX_LAYERED) != 0) result.Add("WS_EX_LAYERED");

            return result.Count > 0 ? string.Join(" | ", result) : "None";
        }

        private string DetermineElementType(string className)
        {
            if (string.IsNullOrEmpty(className))
                return "Unknown";

            // Standard Windows control classes
            var classMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Button", "Button" },
                { "Edit", "Edit/TextBox" },
                { "Static", "Static/Label" },
                { "ListBox", "ListBox" },
                { "ComboBox", "ComboBox" },
                { "ScrollBar", "ScrollBar" },
                { "msctls_progress32", "ProgressBar" },
                { "msctls_trackbar32", "TrackBar/Slider" },
                { "SysTreeView32", "TreeView" },
                { "SysListView32", "ListView" },
                { "SysTabControl32", "TabControl" },
                { "tooltips_class32", "ToolTip" },
                { "SysHeader32", "HeaderControl" },
                { "SysDateTimePick32", "DateTimePicker" },
                { "SysMonthCal32", "MonthCalendar" },
                { "SysIPAddress32", "IPAddressControl" },
                { "RichEdit", "RichEdit" },
                { "RichEdit20W", "RichEdit" },
                { "RichEdit20A", "RichEdit" },
                { "RICHEDIT50W", "RichEdit" },
                { "#32770", "Dialog" },
                { "#32768", "Menu" },
                { "#32769", "Desktop" }
            };

            foreach (var entry in classMap)
            {
                if (className.Contains(entry.Key))
                    return entry.Value;
            }

            // Check for framework-specific classes
            if (className.StartsWith("WindowsForms10"))
                return "WinForms Control";
            if (className.StartsWith("HwndWrapper"))
                return "WPF Window";
            if (className.Contains("Chrome"))
                return "Chrome/Web";
            if (className.Contains("Mozilla"))
                return "Firefox/Web";

            return "Window";
        }
    }
}
