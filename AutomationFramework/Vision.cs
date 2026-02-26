using System.Drawing;
using System.Runtime.InteropServices;
using OpenCvSharp;
using Tesseract;

namespace AutomationFramework;

/// <summary>
/// Provides basic computer vision utilities for desktop automation:
/// screen capture, image template matching, and OCR text lookup.
/// </summary>
public sealed class Vision : IDisposable
{
    public sealed record class Options
    {
        /// <summary>
        /// Folder containing Tesseract language data files (.traineddata).
        /// </summary>
        public string OcrDataPath { get; init; } = Path.Combine(AppContext.BaseDirectory, "tessdata");

        /// <summary>
        /// OCR language code used by Tesseract (for example "eng").
        /// </summary>
        public string OcrLanguage { get; init; } = "eng";

        /// <summary>
        /// The template matching strategy used by OpenCV.
        /// </summary>
        public TemplateMatchModes TemplateMatchMode { get; init; } = TemplateMatchModes.CCoeffNormed;

        internal void Validate()
        {
            if (string.IsNullOrWhiteSpace(OcrDataPath))
            {
                throw new ArgumentException("OCR data path cannot be null or whitespace.", nameof(OcrDataPath));
            }

            if (string.IsNullOrWhiteSpace(OcrLanguage))
            {
                throw new ArgumentException("OCR language cannot be null or whitespace.", nameof(OcrLanguage));
            }
        }
    }

    private readonly Options _options;
    private readonly object _ocrLock = new();
    private TesseractEngine? _ocrEngine;
    private bool _disposed;

    public Vision(Options? options = null)
    {
        _options = options ?? new Options();
        _options.Validate();
    }

    /// <summary>
    /// Captures a screenshot of either the full virtual desktop or a specific region.
    /// </summary>
    public Bitmap CaptureScreenshot(Rectangle? region = null)
    {
        ThrowIfDisposed();

        var captureRegion = NormalizeRegion(region);
        var bitmap = new Bitmap(captureRegion.Width, captureRegion.Height);

        using var graphics = Graphics.FromImage(bitmap);
        graphics.CopyFromScreen(captureRegion.Location, System.Drawing.Point.Empty, captureRegion.Size);
        return bitmap;
    }

    /// <summary>
    /// Captures a screenshot and saves it to the specified path.
    /// </summary>
    public void SaveScreenshot(string outputFilePath, Rectangle? region = null)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFilePath);

        var directory = Path.GetDirectoryName(outputFilePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var screenshot = CaptureScreenshot(region);
        screenshot.Save(outputFilePath);
    }

    /// <summary>
    /// Finds the best match for a template image on screen and returns its confidence score.
    /// </summary>
    public ImageMatchResult? FindImage(string templateImagePath, double minConfidence = 0.8, Rectangle? searchRegion = null)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrWhiteSpace(templateImagePath);

        if (!File.Exists(templateImagePath))
        {
            throw new FileNotFoundException("Template image was not found.", templateImagePath);
        }

        if (minConfidence is < 0 or > 1)
        {
            throw new ArgumentOutOfRangeException(nameof(minConfidence), "minConfidence must be between 0 and 1.");
        }

        var region = NormalizeRegion(searchRegion);

        using var screenshot = CaptureScreenshot(region);
        using var searchMat = ConvertBitmapToMat(screenshot);
        using var templateMat = Cv2.ImRead(templateImagePath, ImreadModes.Color);

        if (templateMat.Empty())
        {
            throw new InvalidOperationException($"Failed to load template image '{templateImagePath}'.");
        }

        if (templateMat.Width > searchMat.Width || templateMat.Height > searchMat.Height)
        {
            throw new ArgumentException("Template image is larger than the search area.", nameof(templateImagePath));
        }

        var resultWidth = searchMat.Width - templateMat.Width + 1;
        var resultHeight = searchMat.Height - templateMat.Height + 1;

        using var resultMat = new Mat(resultHeight, resultWidth, MatType.CV_32FC1);
        Cv2.MatchTemplate(searchMat, templateMat, resultMat, _options.TemplateMatchMode);
        Cv2.MinMaxLoc(resultMat, out var minValue, out var maxValue, out var minLocation, out var maxLocation);

        var (confidence, location) = IsLowerScoreBetter(_options.TemplateMatchMode)
            ? (1 - minValue, minLocation)
            : (maxValue, maxLocation);

        if (confidence < minConfidence)
        {
            return null;
        }

        var bounds = new Rectangle(
            region.Left + location.X,
            region.Top + location.Y,
            templateMat.Width,
            templateMat.Height);

        return new ImageMatchResult(bounds, confidence);
    }

    /// <summary>
    /// Reads all visible text from screen (or a region) and returns OCR words with confidence.
    /// </summary>
    public OcrResult ReadText(Rectangle? searchRegion = null)
    {
        ThrowIfDisposed();

        var region = NormalizeRegion(searchRegion);
        using var screenshot = CaptureScreenshot(region);
        using var page = ProcessOcr(screenshot);
        var fullText = page.GetText() ?? string.Empty;
        var words = ExtractWords(page, region);

        return new OcrResult(fullText, words);
    }

    /// <summary>
    /// Finds OCR words that contain the provided text and meet the minimum confidence.
    /// </summary>
    public IReadOnlyList<TextMatchResult> FindText(
        string text,
        double minConfidence = 0.5,
        Rectangle? searchRegion = null,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrWhiteSpace(text);

        if (minConfidence is < 0 or > 1)
        {
            throw new ArgumentOutOfRangeException(nameof(minConfidence), "minConfidence must be between 0 and 1.");
        }

        var ocrResult = ReadText(searchRegion);
        var matches = ocrResult.Words
            .Where(word => word.Confidence >= minConfidence && word.Text.Contains(text, comparison))
            .Select(word => new TextMatchResult(word.Text, word.Bounds, word.Confidence))
            .ToList();

        return matches;
    }

    private static IReadOnlyList<OcrWord> ExtractWords(Page page, Rectangle region)
    {
        var words = new List<OcrWord>();

        using var iterator = page.GetIterator();
        iterator.Begin();

        do
        {
            var rawText = iterator.GetText(PageIteratorLevel.Word);
            if (string.IsNullOrWhiteSpace(rawText))
            {
                continue;
            }

            if (!iterator.TryGetBoundingBox(PageIteratorLevel.Word, out var box))
            {
                continue;
            }

            var confidence = iterator.GetConfidence(PageIteratorLevel.Word) / 100.0;
            var bounds = new Rectangle(
                region.Left + box.X1,
                region.Top + box.Y1,
                box.Width,
                box.Height);

            words.Add(new OcrWord(rawText.Trim(), bounds, confidence));
        }
        while (iterator.Next(PageIteratorLevel.Word));

        return words;
    }

    private Page ProcessOcr(Bitmap screenshot)
    {
        lock (_ocrLock)
        {
            EnsureOcrEngine();

            using var pix = ConvertBitmapToPix(screenshot);
            return _ocrEngine!.Process(pix);
        }
    }

    private static Mat ConvertBitmapToMat(Bitmap bitmap)
    {
        using var memoryStream = new MemoryStream();
        bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
        var imageBytes = memoryStream.ToArray();
        return Cv2.ImDecode(imageBytes, ImreadModes.Color);
    }

    private static Pix ConvertBitmapToPix(Bitmap bitmap)
    {
        using var memoryStream = new MemoryStream();
        bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
        var imageBytes = memoryStream.ToArray();
        return Pix.LoadFromMemory(imageBytes);
    }

    private void EnsureOcrEngine()
    {
        if (_ocrEngine is not null)
        {
            return;
        }

        if (!Directory.Exists(_options.OcrDataPath))
        {
            throw new DirectoryNotFoundException(
                $"OCR data folder was not found: '{_options.OcrDataPath}'. Add tessdata files (for example eng.traineddata).");
        }

        _ocrEngine = new TesseractEngine(_options.OcrDataPath, _options.OcrLanguage, EngineMode.Default);
    }

    private static Rectangle NormalizeRegion(Rectangle? region)
    {
        if (region is { } selectedRegion)
        {
            if (selectedRegion.Width <= 0 || selectedRegion.Height <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(region), "Region width and height must be greater than zero.");
            }

            return selectedRegion;
        }

        var left = GetSystemMetrics(SystemMetric.XVirtualScreen);
        var top = GetSystemMetrics(SystemMetric.YVirtualScreen);
        var width = GetSystemMetrics(SystemMetric.CxVirtualScreen);
        var height = GetSystemMetrics(SystemMetric.CyVirtualScreen);

        if (width <= 0 || height <= 0)
        {
            throw new InvalidOperationException("Could not determine virtual screen bounds.");
        }

        return new Rectangle(left, top, width, height);
    }

    private static bool IsLowerScoreBetter(TemplateMatchModes mode)
    {
        return mode is TemplateMatchModes.SqDiff or TemplateMatchModes.SqDiffNormed;
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(Vision));
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _ocrEngine?.Dispose();
        _disposed = true;
    }

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(SystemMetric systemMetric);

    private enum SystemMetric
    {
        XVirtualScreen = 76,
        YVirtualScreen = 77,
        CxVirtualScreen = 78,
        CyVirtualScreen = 79
    }
}

public sealed record ImageMatchResult(Rectangle Bounds, double Confidence);

public sealed record OcrWord(string Text, Rectangle Bounds, double Confidence);

public sealed record OcrResult(string Text, IReadOnlyList<OcrWord> Words);

public sealed record TextMatchResult(string Text, Rectangle Bounds, double Confidence);