using System.Numerics;
using AutomationFramework;
using AutomationFramework.Extensions;
using AutomationRunner.Scripting;
using Microsoft.Extensions.Configuration;
using OpenCvSharp;

namespace AutomationRunner.Scripts;

public sealed class VisionDisenchanting : BaseScript
{
    public override string Name => "vision-disenchanting";

    public override string Description => "Disenchants and opens mailbox";

    private AutomationFramework.Cursor _cursor = new();
    private AutomationFramework.Vision _vision = null!;
    private AutomationFramework.Keyboard _keyboard = new();

    private Dictionary<string, Mat> _templates = new Dictionary<string, Mat>();

    protected override Task InitializeAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        _vision = new AutomationFramework.Vision(new AutomationFramework.Vision.Options
        {
            OcrLanguage = "eng",
            OcrDataPath = context.Configuration.GetRequiredSection("OcrDataPath").Value!
        });

        AcquireTemplates
        (
            VisionTemplateFileNames.TSM_OPEN_ALL_MAIL,
            VisionTemplateFileNames.AB_TSM_DESTROY_BTN,
            VisionTemplateFileNames.TSM_DESTROY_NEXT_BTN,
            VisionTemplateFileNames.TSM_CLOSE_BTN
        );

        return Task.CompletedTask;
    }
    
    public override void Dispose()
    {
        if (_vision is null)
        {
            return;
        }

        // Release templates
        foreach (var templateFileName in _templates.Keys)
        {
            _vision.ReleaseTemplate(templateFileName);
        }

        // Dispose vision
        _vision.Dispose();
    }

    protected override async Task RunAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        using var disenchantingCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var disenchantingTask = RunDisenchantingLoopAsync(disenchantingCts.Token);

        try
        {
            while (cancellationToken.IsCancellationRequested == false)
            {
                //wait for disenchanting to free up bag space
                await Task.Delay(TimeSpan.FromSeconds(250).ApplyRandomFactor());

                // Open mailbox
                await _keyboard.PressKeyAsync(AutomationFramework.VirtualKey.F, cancellationToken: cancellationToken);
                await Task.Delay(TimeSpan.FromMilliseconds(500).ApplyRandomFactor(), cancellationToken);

            
                if (await FindAndClickImageTemplateAsync(_templates[VisionTemplateFileNames.TSM_OPEN_ALL_MAIL], bounds => bounds.Padd(40, 2), cancellationToken: cancellationToken) == false)
                {
                    Console.WriteLine("Open all mail button not found.");
                    break;
                }
        
                // wait for all mail to open
                await Task.Delay(TimeSpan.FromSeconds(35).ApplyRandomFactor(), cancellationToken);


                //close mailbox by pressing escape
                await _keyboard.PressKeyAsync(VirtualKey.Escape, cancellationToken: cancellationToken);
            }
        }
        finally
        {
            disenchantingCts.Cancel();

            try
            {
                await disenchantingTask;
            }
            catch (OperationCanceledException)
            {
                // Expected when script stops.
            }
        }
    }

    private async Task RunDisenchantingLoopAsync(CancellationToken cancellationToken)
    {
        while (cancellationToken.IsCancellationRequested == false)
        {
            await _keyboard.PressKeyAsync(AutomationFramework.VirtualKey.D9, cancellationToken: cancellationToken);

            // Wait a random time between key presses to mimic human behavior.
            var delay = TimeSpan.FromSeconds(2.2).ApplyRandomFactor(1f, 1.2f);

            if (delay > TimeSpan.Zero)
            {
                await Task.Delay(delay, cancellationToken);
            }
        }
    }

    private void AcquireTemplates(params string[] filenames)
    {
        ArgumentNullException.ThrowIfNull(filenames);

        foreach (var filename in filenames)
        {
            _templates[filename] = _vision.AcquireTemplate(filename);
        }
    }


    private async Task<bool> FindAndClickImageTemplateAsync(Mat template, Func<Rectangle, Rectangle>? boundsManipulations = null, float confidence = 0.7f, CancellationToken cancellationToken = default)
    {
        // find the template on screen with retries, if not found, return false
        var imageMatch = await Task.RunWithRetry
        (
            (cancellationToken) => _vision.FindImageAsync
            (
                template,
                confidence,
                searchRegion: Screen.PrimaryScreen?.Bounds,
                cancellationToken: cancellationToken
            ),
            successCondition: (result) => result != null,
            maxRetries: 3,
            retryDelay: TimeSpan.FromSeconds(1),
            cancellationToken
        );

        if (imageMatch == null)
        {
            return false;
        }

        // Move cursor to a random point within the target bounds and click
        var targetPos = (boundsManipulations != null ? boundsManipulations(imageMatch.ToGlobalBounds()) : imageMatch.ToGlobalBounds()).GetRandomPointInBounds();
        await _cursor.MoveToAsync(targetPos, cancellationToken: cancellationToken);
        await Task.Delay(TimeSpan.FromMilliseconds(250).ApplyRandomFactor(), cancellationToken);
        await _cursor.ClickAsync(cancellationToken: cancellationToken);

        return true;
    }
}
