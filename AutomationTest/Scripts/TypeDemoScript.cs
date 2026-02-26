using AutomationRunner.Scripting;
using AutomationFramework;
using System.Numerics;

namespace AutomationRunner.Scripts;

public sealed class TypeDemoScript : IAutomationScript
{
    public string Name => "type-demo";

    public string Description => "Waits 3 seconds, then types a demo line into the active window.";

    public async Task ExecuteAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        var cursor = new AutomationFramework.Cursor();

        await cursor.MoveToAsync(new Vector2(0, 900), TimeSpan.FromSeconds(5.3), cancellationToken);


        // var keyboard = new Keyboard();

        // Console.WriteLine("Focus the target app. Typing starts in 3 seconds...");
        // await Task.Delay(TimeSpan.FromSeconds(3), cancellationToken);

        // await keyboard.TypeTextAsync(
        //     "Hello from AutomationFramework script host!",
        //     TimeSpan.FromMilliseconds(35),
        //     cancellationToken);

        // keyboard.PressKey(VirtualKey.Enter);
    }
}
