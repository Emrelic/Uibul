using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Diagnostics;

namespace UIElementInspector.Services
{
    /// <summary>
    /// Service for managing global hotkeys with KeyDown/KeyUp support (shutter mode)
    /// </summary>
    public class HotkeyService : IDisposable
    {
        #region Native Methods

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private const int HOTKEY_ID_START = 9000;
        private const int WM_HOTKEY = 0x0312;
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [Flags]
        private enum ModifierKeys : uint
        {
            None = 0x0000,
            Alt = 0x0001,
            Control = 0x0002,
            Shift = 0x0004,
            Win = 0x0008
        }

        #endregion

        #region Fields

        private readonly Window _window;
        private HwndSource _source;
        private readonly Dictionary<int, Action> _hotkeyActions;
        private readonly List<PendingHotkey> _pendingHotkeys;
        private int _currentHotkeyId;

        // Low-level keyboard hook for KeyDown/KeyUp detection (shutter mode)
        private IntPtr _keyboardHookHandle = IntPtr.Zero;
        private LowLevelKeyboardProc _keyboardProc;
        private readonly Dictionary<uint, ShutterKeyBinding> _shutterBindings;
        private readonly HashSet<uint> _pressedKeys;

        #endregion

        private class PendingHotkey
        {
            public Key Key { get; set; }
            public System.Windows.Input.ModifierKeys Modifiers { get; set; }
            public Action Callback { get; set; }
        }

        private class ShutterKeyBinding
        {
            public Action OnKeyDown { get; set; }
            public Action OnKeyUp { get; set; }
        }

        public HotkeyService(Window window)
        {
            _window = window;
            _hotkeyActions = new Dictionary<int, Action>();
            _pendingHotkeys = new List<PendingHotkey>();
            _shutterBindings = new Dictionary<uint, ShutterKeyBinding>();
            _pressedKeys = new HashSet<uint>();
            _currentHotkeyId = HOTKEY_ID_START;

            // Wait for window to be loaded
            if (_window.IsLoaded)
            {
                Initialize();
            }
            else
            {
                _window.Loaded += Window_Loaded;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Initialize();
            _window.Loaded -= Window_Loaded;
        }

        private void Initialize()
        {
            var helper = new WindowInteropHelper(_window);
            _source = HwndSource.FromHwnd(helper.Handle);
            _source?.AddHook(HwndHook);

            // Register all pending hotkeys
            foreach (var pending in _pendingHotkeys)
            {
                RegisterHotkeyInternal(pending.Key, pending.Modifiers, pending.Callback);
            }
            _pendingHotkeys.Clear();

            // Install low-level keyboard hook for shutter mode
            InstallKeyboardHook();
        }

        private void InstallKeyboardHook()
        {
            if (_keyboardHookHandle != IntPtr.Zero)
                return;

            _keyboardProc = KeyboardHookCallback;
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                _keyboardHookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }

            if (_keyboardHookHandle == IntPtr.Zero)
            {
                Debug.WriteLine("Failed to install keyboard hook");
            }
            else
            {
                Debug.WriteLine("Keyboard hook installed for shutter mode");
            }
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var hookStruct = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                var vkCode = hookStruct.vkCode;
                var msgType = wParam.ToInt32();

                if (_shutterBindings.TryGetValue(vkCode, out var binding))
                {
                    if (msgType == WM_KEYDOWN || msgType == WM_SYSKEYDOWN)
                    {
                        // Only trigger on initial press, not on repeat
                        if (!_pressedKeys.Contains(vkCode))
                        {
                            _pressedKeys.Add(vkCode);
                            _window.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                try
                                {
                                    binding.OnKeyDown?.Invoke();
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"Shutter OnKeyDown error: {ex.Message}");
                                }
                            }));
                        }
                    }
                    else if (msgType == WM_KEYUP || msgType == WM_SYSKEYUP)
                    {
                        _pressedKeys.Remove(vkCode);
                        _window.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            try
                            {
                                binding.OnKeyUp?.Invoke();
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Shutter OnKeyUp error: {ex.Message}");
                            }
                        }));
                    }
                }
            }

            return CallNextHookEx(_keyboardHookHandle, nCode, wParam, lParam);
        }

        /// <summary>
        /// Register a shutter-style hotkey (like a camera shutter - action while held)
        /// </summary>
        /// <param name="key">The key to use as shutter</param>
        /// <param name="onKeyDown">Action when key is pressed down</param>
        /// <param name="onKeyUp">Action when key is released</param>
        public void RegisterShutterKey(Key key, Action onKeyDown, Action onKeyUp)
        {
            var vk = (uint)KeyInterop.VirtualKeyFromKey(key);
            _shutterBindings[vk] = new ShutterKeyBinding
            {
                OnKeyDown = onKeyDown,
                OnKeyUp = onKeyUp
            };
            Debug.WriteLine($"Registered shutter key: {key} (VK: {vk})");
        }

        /// <summary>
        /// Unregister a shutter key
        /// </summary>
        public void UnregisterShutterKey(Key key)
        {
            var vk = (uint)KeyInterop.VirtualKeyFromKey(key);
            _shutterBindings.Remove(vk);
            _pressedKeys.Remove(vk);
            Debug.WriteLine($"Unregistered shutter key: {key}");
        }

        public bool RegisterHotkey(Key key, System.Windows.Input.ModifierKeys modifiers, Action callback)
        {
            if (callback == null) return false;

            // If window is not loaded yet, add to pending list
            if (_source == null)
            {
                _pendingHotkeys.Add(new PendingHotkey
                {
                    Key = key,
                    Modifiers = modifiers,
                    Callback = callback
                });
                return true;
            }

            return RegisterHotkeyInternal(key, modifiers, callback);
        }

        private bool RegisterHotkeyInternal(Key key, System.Windows.Input.ModifierKeys modifiers, Action callback)
        {
            if (callback == null) return false;

            var vk = (uint)KeyInterop.VirtualKeyFromKey(key);
            var modifierFlags = ConvertModifiers(modifiers);

            var id = _currentHotkeyId++;
            _hotkeyActions[id] = callback;

            var helper = new WindowInteropHelper(_window);
            var result = RegisterHotKey(helper.Handle, id, (uint)modifierFlags, vk);

            if (!result)
            {
                _hotkeyActions.Remove(id);
                System.Diagnostics.Debug.WriteLine($"Failed to register hotkey: {key} with modifiers {modifiers}");
                return false;
            }

            System.Diagnostics.Debug.WriteLine($"Successfully registered global hotkey: {key} with modifiers {modifiers}");
            return true;
        }

        public bool RegisterHotkey(Key key, System.Windows.Input.ModifierKeys modifiers,
            EventHandler<RoutedEventArgs> handler)
        {
            return RegisterHotkey(key, modifiers, () => handler?.Invoke(this, new RoutedEventArgs()));
        }

        private ModifierKeys ConvertModifiers(System.Windows.Input.ModifierKeys modifiers)
        {
            var result = ModifierKeys.None;

            if ((modifiers & System.Windows.Input.ModifierKeys.Alt) != 0)
                result |= ModifierKeys.Alt;
            if ((modifiers & System.Windows.Input.ModifierKeys.Control) != 0)
                result |= ModifierKeys.Control;
            if ((modifiers & System.Windows.Input.ModifierKeys.Shift) != 0)
                result |= ModifierKeys.Shift;
            if ((modifiers & System.Windows.Input.ModifierKeys.Windows) != 0)
                result |= ModifierKeys.Win;

            return result;
        }

        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY)
            {
                int hotkeyId = wParam.ToInt32();

                if (_hotkeyActions.TryGetValue(hotkeyId, out var action))
                {
                    action?.Invoke();
                    handled = true;
                }
            }

            return IntPtr.Zero;
        }

        public void UnregisterAll()
        {
            var helper = new WindowInteropHelper(_window);

            foreach (var id in _hotkeyActions.Keys)
            {
                UnregisterHotKey(helper.Handle, id);
            }

            _hotkeyActions.Clear();
            _shutterBindings.Clear();
            _pressedKeys.Clear();
        }

        public void Dispose()
        {
            UnregisterAll();

            // Remove keyboard hook
            if (_keyboardHookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_keyboardHookHandle);
                _keyboardHookHandle = IntPtr.Zero;
                Debug.WriteLine("Keyboard hook removed");
            }

            _source?.RemoveHook(HwndHook);
            _source?.Dispose();
        }
    }
}