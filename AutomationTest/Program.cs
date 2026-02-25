using System.Numerics;
using System.Runtime.InteropServices;
using AutomationFramework;
using GlobalHotkeyManager;

namespace AutomationTest
{
    internal class Program
    {
        [DllImport("kernel32.dll")]
        private static extern nint GetConsoleWindow();

        static CancellationTokenSource cts = new();

        static async Task<int> Main(string[] args)
        {
            HotkeyManager hotKeyManager = new HotkeyManager();
            hotKeyManager.HotkeyPressed += OnHotKeyPressed;

            hotKeyManager.RegisterHotkey(Keys.C, KeyModifiers.Control);
            hotKeyManager.RegisterHotkey(Keys.Q, KeyModifiers.Alt);

            AutomationFramework.Cursor cursor = new();
           
            var consoleWindow = GetConsoleWindow();
            var launchScreen = consoleWindow != nint.Zero
                ? Screen.FromHandle(consoleWindow)
                : Screen.PrimaryScreen ?? Screen.AllScreens[0];

            var monitorBounds = launchScreen.Bounds;
            var currentPos = new Vector2(
                monitorBounds.Left + (monitorBounds.Width / 2f),
                monitorBounds.Top + (monitorBounds.Height / 2f));

            try
            {
                while (!cts.IsCancellationRequested)
                {
                    float x = Random.Shared.Next(monitorBounds.Left, monitorBounds.Right);
                    float y = Random.Shared.Next(monitorBounds.Top, monitorBounds.Bottom);
                    var targetPos = new Vector2(x, y);

                    await cursor.MoveToAsync(
                        targetPos,
                        TimeSpan.FromMilliseconds(Random.Shared.Next(600, 1800)),
                        cts.Token);

                    currentPos = targetPos;

                    await Task.Delay(Random.Shared.Next(80, 250), cts.Token);
                }
            }
            catch (OperationCanceledException)
            {
                
            }

            hotKeyManager.Dispose();

            return 0;
        }

        private static void OnHotKeyPressed(object? sender, HotkeyEventArgs e)
        {
            Console.WriteLine($"Hotkey pressed: {e.Modifiers} + {e.Key}");

            if (e.Key == Keys.Q && e.Modifiers.HasFlag(KeyModifiers.Alt))
            {
                Console.WriteLine($"Current cursor position: {System.Windows.Forms.Cursor.Position}");
            }

            if (e.Key == Keys.C && e.Modifiers.HasFlag(KeyModifiers.Control))
            {
                cts.Cancel();
            }
        }
    }
}