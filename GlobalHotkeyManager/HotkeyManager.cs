using System.ComponentModel;
using System.Runtime.InteropServices;

namespace GlobalHotkeyManager
{
    public sealed class HotkeyManager
    {
        private readonly ManualResetEventSlim _windowReadyEvent = new(false);
        private readonly HashSet<int> _registeredHotkeyIds = [];
        private readonly Thread _messageLoopThread;

        private MessageWindow? _window;
        private IntPtr _windowHandle;
        private int _id;
        private bool _disposed;

        public event EventHandler<HotkeyEventArgs>? HotkeyPressed;

        public HotkeyManager()
        {
            _messageLoopThread = new Thread(() => Application.Run(new MessageWindow(this)))
            {
                Name = "HotkeyMessageLoopThread",
                IsBackground = true
            };

            _messageLoopThread.SetApartmentState(ApartmentState.STA);

            _messageLoopThread.Start();
        }

        public int RegisterHotkey(Keys key, KeyModifiers modifiers)
        {
            ThrowIfDisposed();
            _windowReadyEvent.Wait();

            var id = Interlocked.Increment(ref _id);

            InvokeOnWindowThread(() =>
            {
                if (!RegisterHotkeyNative(_windowHandle, id, (uint)modifiers, (uint)key))
                {
                    var errorCode = Marshal.GetLastWin32Error();
                    var win32Exception = new Win32Exception(errorCode);
                    throw new InvalidOperationException(
                        $"Failed to register hotkey '{modifiers} + {key}' (Win32 error {errorCode}: {win32Exception.Message}). The key combination may already be in use by another application.",
                        win32Exception);
                }

                _registeredHotkeyIds.Add(id);
            });

            return id;
        }

        public void UnregisterHotkey(int id)
        {
            if (_disposed)
            {
                return;
            }

            _windowReadyEvent.Wait();

            InvokeOnWindowThread(() =>
            {
                UnregisterHotkeyNative(_windowHandle, id);
                _registeredHotkeyIds.Remove(id);
            });
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            if (_windowReadyEvent.IsSet && _window is not null)
            {
                InvokeOnWindowThread(() =>
                {
                    foreach (var id in _registeredHotkeyIds.ToArray())
                    {
                        UnregisterHotkeyNative(_windowHandle, id);
                    }

                    _registeredHotkeyIds.Clear();
                    _window.Close();
                    Application.ExitThread();
                });
            }

            _windowReadyEvent.Dispose();
            GC.SuppressFinalize(this);
        }

        private void InvokeOnWindowThread(Action action)
        {
            if (_window is null)
            {
                throw new InvalidOperationException("Hotkey message window is not available.");
            }

            _window.Invoke(action);
        }

        private void InitializeWindow(MessageWindow window)
        {
            _window = window;
            _windowHandle = window.Handle;
            _windowReadyEvent.Set();
        }

        private void RaiseHotkeyPressed(HotkeyEventArgs args)
        {
            HotkeyPressed?.Invoke(this, args);
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(HotkeyManager));
            }
        }

        private sealed class MessageWindow : Form
        {
            private const int WmHotkey = 0x312;
            private readonly HotkeyManager _owner;

            public MessageWindow(HotkeyManager owner)
            {
                _owner = owner;
                _owner.InitializeWindow(this);
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WmHotkey)
                {
                    Console.WriteLine($"WParam: {m.WParam}, LParam: {m.LParam}");
                    _owner.RaiseHotkeyPressed(new HotkeyEventArgs(m.LParam));
                }

                base.WndProc(ref m);
            }

            protected override void SetVisibleCore(bool value)
            {
                base.SetVisibleCore(false);
            }
        }

        [DllImport("user32", EntryPoint = "RegisterHotKey", SetLastError = true)]
        private static extern bool RegisterHotkeyNative(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32", EntryPoint = "UnregisterHotKey", SetLastError = true)]
        private static extern bool UnregisterHotkeyNative(IntPtr hWnd, int id);
    }

    public sealed class HotkeyEventArgs : EventArgs
    {
        public Keys Key { get; }
        public KeyModifiers Modifiers { get; }

        public HotkeyEventArgs(Keys key, KeyModifiers modifiers)
        {
            Key = key;
            Modifiers = modifiers;
        }

        public HotkeyEventArgs(IntPtr hotkeyParam)
        {
            uint param = (uint)hotkeyParam.ToInt64();
            Key = (Keys)((param & 0xffff0000) >> 16);
            Modifiers = (KeyModifiers)(param & 0x0000ffff);
        }
    }

    [Flags]
    public enum KeyModifiers : uint
    {
        None = 0x0000,
        Alt = 0x0001,
        Control = 0x0002,
        Shift = 0x0004,
        Windows = 0x0008,
        NoRepeat = 0x4000
    }
}