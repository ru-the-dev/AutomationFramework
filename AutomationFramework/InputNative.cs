using System.Runtime.InteropServices;

namespace AutomationFramework;

internal static class InputNative
{
    private const uint InputMouse = 0;
    private const uint InputKeyboard = 1;

    private const uint MouseEventFMove = 0x0001;
    private const uint MouseEventFLeftDown = 0x0002;
    private const uint MouseEventFLeftUp = 0x0004;
    private const uint MouseEventFRightDown = 0x0008;
    private const uint MouseEventFRightUp = 0x0010;
    private const uint MouseEventFMiddleDown = 0x0020;
    private const uint MouseEventFMiddleUp = 0x0040;
    private const uint MouseEventFAbsolute = 0x8000;
    private const uint MouseEventFVirtualDesk = 0x4000;

    private const uint KeyEventFKeyUp = 0x0002;
    private const uint KeyEventFUnicode = 0x0004;

    private const int SmXVirtualScreen = 76;
    private const int SmYVirtualScreen = 77;
    private const int SmCxVirtualScreen = 78;
    private const int SmCyVirtualScreen = 79;

    internal static void SendMouseAbsolute(int x, int y)
    {
        var virtualLeft = GetSystemMetrics(SmXVirtualScreen);
        var virtualTop = GetSystemMetrics(SmYVirtualScreen);
        var virtualWidth = Math.Max(1, GetSystemMetrics(SmCxVirtualScreen));
        var virtualHeight = Math.Max(1, GetSystemMetrics(SmCyVirtualScreen));

        var normalizedX = NormalizeAbsoluteCoordinate(x, virtualLeft, virtualWidth);
        var normalizedY = NormalizeAbsoluteCoordinate(y, virtualTop, virtualHeight);

        var input = new Input
        {
            Type = InputMouse,
            Data = new InputUnion
            {
                Mouse = new MouseInput
                {
                    Dx = normalizedX,
                    Dy = normalizedY,
                    MouseData = 0,
                    DwFlags = MouseEventFMove | MouseEventFAbsolute | MouseEventFVirtualDesk,
                    Time = 0,
                    DwExtraInfo = IntPtr.Zero
                }
            }
        };

        SendInputOrThrow(input);
    }

    internal static void SendMouseButtonDown(MouseButton button)
    {
        var flags = button switch
        {
            MouseButton.Left => MouseEventFLeftDown,
            MouseButton.Right => MouseEventFRightDown,
            MouseButton.Middle => MouseEventFMiddleDown,
            _ => throw new ArgumentOutOfRangeException(nameof(button), button, "Unsupported mouse button.")
        };

        SendMouseFlag(flags);
    }

    internal static void SendMouseButtonUp(MouseButton button)
    {
        var flags = button switch
        {
            MouseButton.Left => MouseEventFLeftUp,
            MouseButton.Right => MouseEventFRightUp,
            MouseButton.Middle => MouseEventFMiddleUp,
            _ => throw new ArgumentOutOfRangeException(nameof(button), button, "Unsupported mouse button.")
        };

        SendMouseFlag(flags);
    }

    internal static void SendKeyDown(ushort virtualKey)
    {
        var input = new Input
        {
            Type = InputKeyboard,
            Data = new InputUnion
            {
                Keyboard = new KeyboardInput
                {
                    WVk = virtualKey,
                    WScan = 0,
                    DwFlags = 0,
                    Time = 0,
                    DwExtraInfo = IntPtr.Zero
                }
            }
        };

        SendInputOrThrow(input);
    }

    internal static void SendKeyUp(ushort virtualKey)
    {
        var input = new Input
        {
            Type = InputKeyboard,
            Data = new InputUnion
            {
                Keyboard = new KeyboardInput
                {
                    WVk = virtualKey,
                    WScan = 0,
                    DwFlags = KeyEventFKeyUp,
                    Time = 0,
                    DwExtraInfo = IntPtr.Zero
                }
            }
        };

        SendInputOrThrow(input);
    }

    internal static void SendUnicodeChar(char character)
    {
        var down = new Input
        {
            Type = InputKeyboard,
            Data = new InputUnion
            {
                Keyboard = new KeyboardInput
                {
                    WVk = 0,
                    WScan = character,
                    DwFlags = KeyEventFUnicode,
                    Time = 0,
                    DwExtraInfo = IntPtr.Zero
                }
            }
        };

        var up = new Input
        {
            Type = InputKeyboard,
            Data = new InputUnion
            {
                Keyboard = new KeyboardInput
                {
                    WVk = 0,
                    WScan = character,
                    DwFlags = KeyEventFUnicode | KeyEventFKeyUp,
                    Time = 0,
                    DwExtraInfo = IntPtr.Zero
                }
            }
        };

        SendInputOrThrow(down, up);
    }

    private static int NormalizeAbsoluteCoordinate(int value, int axisOffset, int axisLength)
    {
        if (axisLength <= 1)
        {
            return 0;
        }

        var clamped = Math.Clamp(value - axisOffset, 0, axisLength - 1);
        return (int)Math.Round(clamped * 65535.0 / (axisLength - 1));
    }

    private static void SendMouseFlag(uint flags)
    {
        var input = new Input
        {
            Type = InputMouse,
            Data = new InputUnion
            {
                Mouse = new MouseInput
                {
                    Dx = 0,
                    Dy = 0,
                    MouseData = 0,
                    DwFlags = flags,
                    Time = 0,
                    DwExtraInfo = IntPtr.Zero
                }
            }
        };

        SendInputOrThrow(input);
    }

    private static void SendInputOrThrow(params Input[] inputs)
    {
        var sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<Input>());
        if (sent != inputs.Length)
        {
            throw new InvalidOperationException($"SendInput failed. Win32 error: {Marshal.GetLastWin32Error()}.");
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int SendInput(uint cInputs, [MarshalAs(UnmanagedType.LPArray), In] Input[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    [StructLayout(LayoutKind.Sequential)]
    private struct Input
    {
        public uint Type;
        public InputUnion Data;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MouseInput Mouse;

        [FieldOffset(0)]
        public KeyboardInput Keyboard;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput
    {
        public int Dx;
        public int Dy;
        public uint MouseData;
        public uint DwFlags;
        public uint Time;
        public IntPtr DwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardInput
    {
        public ushort WVk;
        public ushort WScan;
        public uint DwFlags;
        public uint Time;
        public IntPtr DwExtraInfo;
    }
}