using System.Numerics;
using AutomationRunner.Scripting;
using AutomationFramework;
using Microsoft.Extensions.Configuration;

namespace AutomationRunner.Scripts;

public sealed class VisionTestScript : IAutomationScript
{
    public string Name => "vision-test";

    public string Description => "Finds a template, manually converts local match bounds to global coordinates, then clicks.";

    public async Task ExecuteAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        using var vision = new Vision(new Vision.Options
        {
            OcrLanguage = "eng",
            OcrDataPath = context.Configuration.GetRequiredSection("OcrDataPath").Value!
        });

        var cursor = new AutomationFramework.Cursor();

        var windowsKeyMatchResult = vision.FindImage("windows-button.png", 0.8);
        if (windowsKeyMatchResult is null)
        {
            Console.WriteLine("Failed to find windows key on screen.");
            return;
        }

        var globalBounds = windowsKeyMatchResult.ToGlobalBounds();
        
        var target = new Vector2(
            globalBounds.Left + (globalBounds.Width / 2f),
            globalBounds.Top + (globalBounds.Height / 2f));

        await cursor.MoveToAsync(target, TimeSpan.FromMilliseconds(450), cancellationToken);
        cursor.Click();

        Console.WriteLine($"Clicked at X={target.X:0}, Y={target.Y:0}");
    }
}
