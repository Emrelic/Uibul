using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;

namespace UIElementInspector.Services
{
    /// <summary>
    /// Service for global mouse hook to capture mouse events
    /// </summary>
    public class MouseHookService : IDisposable
    {
        #region Native Methods and Types

        private const int WH_MOUSE_LL = 14;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn,
            IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private enum MouseMessages
        {
            WM_LBUTTONDOWN = 0x0201,
            WM_LBUTTONUP = 0x0202,
            WM_MOUSEMOVE = 0x0200,
            WM_MOUSEWHEEL = 0x020A,
            WM_RBUTTONDOWN = 0x0204,
            WM_RBUTTONUP = 0x0205
        }

        #endregion

        #region Fields and Properties

        private LowLevelMouseProc _proc;
        private IntPtr _hookID = IntPtr.Zero;
        private bool _isHookActive;
        private System.Windows.Point _lastPosition;
        private DateTime _lastMoveTime;
        private readonly TimeSpan _moveThreshold = TimeSpan.FromMilliseconds(50);

        public bool IsHookActive => _isHookActive;

        #endregion

        #region Events

        public event EventHandler<System.Windows.Point> MouseMove;
        public event EventHandler<System.Windows.Point> MouseClick;
        public event EventHandler<System.Windows.Point> MouseRightClick;
        public event EventHandler<int> MouseWheel;

        #endregion

        public MouseHookService()
        {
            _proc = HookCallback;
        }

        public void StartHook()
        {
            if (_isHookActive) return;

            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                _hookID = SetWindowsHookEx(WH_MOUSE_LL, _proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }

            _isHookActive = true;
            Debug.WriteLine("Mouse hook started");
        }

        public void StopHook()
        {
            if (!_isHookActive) return;

            UnhookWindowsHookEx(_hookID);
            _hookID = IntPtr.Zero;
            _isHookActive = false;
            Debug.WriteLine("Mouse hook stopped");
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0)
                {
                    MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    System.Windows.Point point = new System.Windows.Point(hookStruct.pt.x, hookStruct.pt.y);

                    switch ((MouseMessages)wParam)
                    {
                        case MouseMessages.WM_MOUSEMOVE:
                            HandleMouseMove(point);
                            break;

                        case MouseMessages.WM_LBUTTONDOWN:
                            MouseClick?.Invoke(this, point);
                            break;

                        case MouseMessages.WM_RBUTTONDOWN:
                            MouseRightClick?.Invoke(this, point);
                            break;

                        case MouseMessages.WM_MOUSEWHEEL:
                            int delta = (int)(hookStruct.mouseData >> 16);
                            MouseWheel?.Invoke(this, delta);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Mouse hook error: {ex.Message}");
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private void HandleMouseMove(System.Windows.Point point)
        {
            // Throttle mouse move events to avoid overwhelming the system
            var now = DateTime.UtcNow;

            if (point != _lastPosition && (now - _lastMoveTime) > _moveThreshold)
            {
                _lastPosition = point;
                _lastMoveTime = now;
                MouseMove?.Invoke(this, point);
            }
        }

        public void Dispose()
        {
            StopHook();
        }
    }
}