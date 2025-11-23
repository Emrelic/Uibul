using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace UIElementInspector.Services
{
    /// <summary>
    /// Service for managing global hotkeys
    /// </summary>
    public class HotkeyService : IDisposable
    {
        #region Native Methods

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int HOTKEY_ID_START = 9000;
        private const int WM_HOTKEY = 0x0312;

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

        #endregion

        private class PendingHotkey
        {
            public Key Key { get; set; }
            public System.Windows.Input.ModifierKeys Modifiers { get; set; }
            public Action Callback { get; set; }
        }

        public HotkeyService(Window window)
        {
            _window = window;
            _hotkeyActions = new Dictionary<int, Action>();
            _pendingHotkeys = new List<PendingHotkey>();
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
        }

        public void Dispose()
        {
            UnregisterAll();
            _source?.RemoveHook(HwndHook);
            _source?.Dispose();
        }
    }
}