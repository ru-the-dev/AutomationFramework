using AutomationTest.Scripting;
using AutomationFramework;
using System.Numerics;

namespace AutomationTest.Scripts;

public sealed class VisionDemoScript : IAutomationScript
{
    public string Name => "vision-demo";

    public string Description => "Finds a template image/OCR text and shows manual local-to-global cursor targeting.";

    public async Task ExecuteAsync(CancellationToken cancellationToken)
    {
        var templatePath = Path.Combine(AppContext.BaseDirectory, "assets", "template.png");

        using var vision = new Vision(new Vision.Options
        {
            OcrDataPath = Path.Combine(AppContext.BaseDirectory, "tessdata"),
            OcrLanguage = "eng"
        });

        var cursor = new AutomationFramework.Cursor();

        if (File.Exists(templatePath))
        {
            var imageMatch = vision.FindImage(templatePath, minConfidence: 0.8);
            if (imageMatch is null)
            {
                Console.WriteLine("Template image was not found on screen above confidence 0.80.");
            }
            else
            {
                var globalBounds = imageMatch.ToGlobalBounds();
                var centerX = globalBounds.Left + (globalBounds.Width / 2f);
                var centerY = globalBounds.Top + (globalBounds.Height / 2f);

                Console.WriteLine(
                    $"Template match: confidence={imageMatch.Confidence:P1}, local={imageMatch.Bounds.X},{imageMatch.Bounds.Y},{imageMatch.Bounds.Width},{imageMatch.Bounds.Height}, global={globalBounds.X},{globalBounds.Y},{globalBounds.Width},{globalBounds.Height}");

                // Manual conversion is explicit in the script: local match bounds -> global screen point.
                Console.Write("Move cursor to template center and click? (y/N): ");
                var shouldClick = Console.ReadLine();
                if (string.Equals(shouldClick, "y", StringComparison.OrdinalIgnoreCase))
                {
                    await cursor.MoveToAsync(new Vector2(centerX, centerY), TimeSpan.FromMilliseconds(450), cancellationToken);
                    cursor.Click();
                    Console.WriteLine("Clicked template center.");
                }
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
                    var global = match.ToGlobalBounds();
                    Console.WriteLine(
                        $"- '{match.Text}' conf={match.Confidence:P1} local={match.Bounds.X},{match.Bounds.Y},{match.Bounds.Width},{match.Bounds.Height} global={global.X},{global.Y},{global.Width},{global.Height}");
                }
            }
        }
        catch (DirectoryNotFoundException ex)
        {
            Console.WriteLine(ex.Message);
            Console.WriteLine("OCR setup: copy tessdata (for example eng.traineddata) to AutomationTest/bin/Debug/net10.0-windows/tessdata.");
        }

    }
}
