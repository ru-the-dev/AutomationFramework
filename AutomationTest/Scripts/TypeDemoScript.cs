using AutomationTest.Scripting;
using AutomationFramework;

namespace AutomationTest.Scripts;

public sealed class TypeDemoScript : IAutomationScript
{
    public string Name => "type-demo";

    public string Description => "Waits 3 seconds, then types a demo line into the active window.";

    public async Task ExecuteAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        Console.WriteLine("Focus the target app. Typing starts in 3 seconds...");
        await Task.Delay(TimeSpan.FromSeconds(3), cancellationToken);

        await context.Keyboard.TypeTextAsync(
            "Hello from AutomationFramework script host!",
            TimeSpan.FromMilliseconds(35),
            cancellationToken);

        context.Keyboard.PressKey(VirtualKey.Enter);
    }
}
