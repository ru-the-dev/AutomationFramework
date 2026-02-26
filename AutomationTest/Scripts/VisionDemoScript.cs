using AutomationTest.Scripting;
using AutomationFramework;

namespace AutomationTest.Scripts;

public sealed class VisionDemoScript : IAutomationScript
{
    public string Name => "vision-demo";

    public string Description => "Captures a screenshot, optionally matches assets/template.png, and runs OCR text lookup.";

    public Task ExecuteAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var outputDirectory = Path.Combine(AppContext.BaseDirectory, "artifacts");
        Directory.CreateDirectory(outputDirectory);

        var screenshotPath = Path.Combine(outputDirectory, $"screenshot-{DateTime.Now:yyyyMMdd-HHmmss}.png");
        var templatePath = Path.Combine(AppContext.BaseDirectory, "assets", "template.png");

        using var vision = new Vision(new Vision.Options
        {
            OcrDataPath = Path.Combine(AppContext.BaseDirectory, "tessdata"),
            OcrLanguage = "eng"
        });

        vision.SaveScreenshot(screenshotPath);
        Console.WriteLine($"Saved screenshot: {screenshotPath}");

        if (File.Exists(templatePath))
        {
            var imageMatch = vision.FindImage(templatePath, minConfidence: 0.8);
            if (imageMatch is null)
            {
                Console.WriteLine("Template image was not found on screen above confidence 0.80.");
            }
            else
            {
                Console.WriteLine(
                    $"Template match: confidence={imageMatch.Confidence:P1}, bounds={imageMatch.Bounds.X},{imageMatch.Bounds.Y},{imageMatch.Bounds.Width},{imageMatch.Bounds.Height}");
            }
        }
        else
        {
            Console.WriteLine($"Template skipped. Add a template image at: {templatePath}");
        }

        try
        {
            var ocrResult = vision.ReadText();
            var preview = ocrResult.Text.ReplaceLineEndings(" ").Trim();
            if (preview.Length > 180)
            {
                preview = preview[..180] + "...";
            }

            Console.WriteLine(string.IsNullOrWhiteSpace(preview)
                ? "OCR extracted no text from the current screen."
                : $"OCR preview: {preview}");

            Console.Write("Enter text to search on screen (or press Enter to skip): ");
            var query = Console.ReadLine();

            if (!string.IsNullOrWhiteSpace(query))
            {
                var matches = vision.FindText(query, minConfidence: 0.55);
                Console.WriteLine($"Found {matches.Count} matching OCR word(s) for '{query}'.");

                foreach (var match in matches.Take(5))
                {
                    Console.WriteLine(
                        $"- '{match.Text}' conf={match.Confidence:P1} at {match.Bounds.X},{match.Bounds.Y},{match.Bounds.Width},{match.Bounds.Height}");
                }
            }
        }
        catch (DirectoryNotFoundException ex)
        {
            Console.WriteLine(ex.Message);
            Console.WriteLine("OCR setup: copy tessdata (for example eng.traineddata) to AutomationTest/bin/Debug/net10.0-windows/tessdata.");
        }

        return Task.CompletedTask;
    }
}
