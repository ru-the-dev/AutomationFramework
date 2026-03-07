

using System.Data.Common;
using System.Numerics;
using AutomationFramework;
using AutomationFramework.Extensions;
using AutomationRunner.Scripting;
using Microsoft.Extensions.Configuration;
using OpenCvSharp;

namespace AutomationRunner.Scripts;

public sealed class VisionTest : BaseScript
{
    public override string Name => "vision-test";

    public override string Description => "A script to test vision functionality.";

    AutomationFramework.Cursor _cursor = new();
    AutomationFramework.Vision _vision = null!;

    Dictionary<string, Mat> _templates = new Dictionary<string, Mat>();


    protected override Task InitializeAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        _vision = new AutomationFramework.Vision(new AutomationFramework.Vision.Options
        {
            OcrLanguage = "eng",
            OcrDataPath = context.Configuration.GetRequiredSection("OcrDataPath").Value!,
            TemplateMatchMode = TemplateMatchModes.CCoeffNormed
        });

        // Load templates
        AcquireTemplates
        (
            VisionTemplateFileNames.TailoringButton,
            VisionTemplateFileNames.TsmMaxButton,
            VisionTemplateFileNames.TsmCraftButton,
            VisionTemplateFileNames.TsmCloseButton,
            VisionTemplateFileNames.GildedTradersBrutosaur,
            VisionTemplateFileNames.TargetMailButton,
            VisionTemplateFileNames.TsmMailboxGroupsButton,
            VisionTemplateFileNames.TsmMailSelectedGroupsButton
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
        var res = await _vision.FindImageByEdgePyramidAsync
        (
            _templates[VisionTemplateFileNames.TailoringButton],
            minConfidence: 0.80,
            minColorCorrelation: 0.25,
            attempts: 5,
            searchRegion: Screen.PrimaryScreen?.Bounds,
            cancellationToken: cancellationToken
        );

        if (res == null)
        {
            System.Console.WriteLine("Image not found.");
            return;
        }

        var target = res.ToGlobalBounds().Center();
        Console.WriteLine($"Found match with confidence {res.Confidence:F3} at {res.ToGlobalBounds()}");
        await _cursor.MoveToAsync(target, cancellationToken: cancellationToken);

    }   

    private void AcquireTemplates(params string[] filenames)
    {
        ArgumentNullException.ThrowIfNull(filenames);

        foreach (var filename in filenames)
        {
            _templates[filename] = _vision.AcquireTemplate(filename);
        }
    }

    
}