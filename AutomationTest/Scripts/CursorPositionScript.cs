using AutomationTest.Scripting;

namespace AutomationTest.Scripts;

public sealed class CursorPositionScript : IAutomationScript
{
    public string Name => "cursor-position";

    public string Description => "Prints the current cursor position once per second for 5 seconds.";

    public async Task ExecuteAsync(CancellationToken cancellationToken)
    {
        var cursor = new AutomationFramework.Cursor(); 

        for (var second = 1; second <= 5; second++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var position = cursor.GetCurrentPosition();
            Console.WriteLine($"[{second}] X={position.X:0}, Y={position.Y:0}");
            await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
        }
    }
}
