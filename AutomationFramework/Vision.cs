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
    /// <summary>
    /// Configuration for <see cref="Vision"/>.
    /// </summary>
    public sealed record class Options
    {
        /// <summary>
        /// Folder containing Tesseract language data files (.traineddata).
        /// </summary>
        public string OcrDataPath { get; init; } = Path.Combine(AppContext.BaseDirectory, "tessdata");

        /// <summary>
        /// Folder containing template images for matching. Defaults to "Assets/Templates" subfolder of the current directory.
        /// </summary>
        public string TemplatePath { get; init; } = Path.Combine(AppContext.BaseDirectory, "Assets", "Templates");


        /// <summary>
        /// OCR language code used by Tesseract (for example "eng").
        /// </summary>
        public string OcrLanguage { get; init; } = "eng";
        

        /// <summary>
        /// The template matching strategy used by OpenCV.
        /// </summary>
        public TemplateMatchModes TemplateMatchMode { get; init; } = TemplateMatchModes.CCoeffNormed;

        /// <summary>
        /// Validates option values before the engine is used.
        /// </summary>
        /// <exception cref="ArgumentException">
        /// Thrown when required values are invalid.
        /// </exception>
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

            if (string.IsNullOrWhiteSpace(TemplatePath))
            {
                throw new ArgumentException("Template path cannot be null or whitespace.", nameof(TemplatePath));
            }
        }
    }

    private readonly Options _options;
    private readonly object _ocrLock = new();
    private TesseractEngine? _ocrEngine;
    private bool _disposed;

    /// <summary>
    /// Creates a new vision helper with optional custom settings.
    /// </summary>
    /// <param name="options">Optional vision settings. Defaults are used when omitted.</param>
    public Vision(Options? options = null)
    {
        _options = options ?? new Options();
        _options.Validate();
    }

    /// <summary>
    /// Captures a screenshot of either the full virtual desktop or a specific region.
    /// </summary>
    /// <param name="region">
    /// Optional screen region to capture. When null, the full virtual desktop is captured.
    /// </param>
    /// <returns>A bitmap containing the captured pixels.</returns>
    /// <exception cref="ObjectDisposedException">Thrown when this instance is disposed.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when region size is invalid.</exception>
    /// <exception cref="InvalidOperationException">Thrown when virtual screen bounds cannot be read.</exception>
    public Bitmap CaptureScreenshot(Rectangle? region = null)
    {
        ThrowIfDisposed();

        // Normalize to caller-provided region, or configured default scope.
        var captureRegion = NormalizeRegion(region);
        var bitmap = new Bitmap(captureRegion.Width, captureRegion.Height);

        // Copy pixels from screen into a top-left anchored bitmap.
        using var graphics = Graphics.FromImage(bitmap);
        graphics.CopyFromScreen(captureRegion.Location, System.Drawing.Point.Empty, captureRegion.Size);
        return bitmap;
    }

    /// <summary>
    /// Captures a screenshot and saves it to the specified path.
    /// </summary>
    /// <param name="outputFilePath">Destination file path (extension determines image format).</param>
    /// <param name="region">
    /// Optional screen region to capture. When null, the full virtual desktop is captured.
    /// </param>
    /// <exception cref="ObjectDisposedException">Thrown when this instance is disposed.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="outputFilePath"/> is invalid.</exception>
    public void SaveScreenshot(string outputFilePath, Rectangle? region = null)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFilePath);

        // Ensure target folder exists before writing the file.
        var directory = Path.GetDirectoryName(outputFilePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var screenshot = CaptureScreenshot(region);
        screenshot.Save(outputFilePath);
    }


    public Mat AcquireTemplate(string templateFileName)
    {
        var templateFilePath = Path.Combine(_options.TemplatePath, templateFileName);
        return VisionTemplateResourceManager.Acquire(templateFilePath);
    }

    public void ReleaseTemplate(string templateFileName)
    {
        var templateFilePath = Path.Combine(_options.TemplatePath, templateFileName);
        VisionTemplateResourceManager.Release(templateFilePath);
    }

    /// <summary>
    /// Finds the best match for a template image on screen and returns its confidence score.
    /// </summary>
    /// <param name="templateImagePath">Path to the template image to locate.</param>
    /// <param name="minConfidence">
    /// Minimum accepted confidence in the range [0..1]. Higher values are stricter.
    /// </param>
    /// <param name="searchRegion">
    /// Optional region to search. When null, the full virtual desktop is searched.
    /// </param>
    /// <returns>
    /// The best image match if confidence is at least <paramref name="minConfidence"/>; otherwise null.
    /// </returns>
    /// <exception cref="ObjectDisposedException">Thrown when this instance is disposed.</exception>
    /// <exception cref="ArgumentException">Thrown when inputs are invalid.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when confidence is outside [0..1].</exception>
    /// <exception cref="FileNotFoundException">Thrown when the template file does not exist.</exception>
    /// <exception cref="InvalidOperationException">Thrown when template loading fails.</exception>
    public async Task<ImageMatchResult?> FindImageAsync
    (   
        string templateImage, 
        double minConfidence = 0.8, 
        Rectangle? searchRegion = null, 
        int attempts = 10,
        int retryDelayMs = 250,
        float scaleStep = 0.05f,
        float maxScale = 2.0f,
        float minScale = 0.25f,
        CancellationToken cancellationToken = default
    )
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrWhiteSpace(templateImage);

        var templateImagePath = Path.Combine(_options.TemplatePath, templateImage); 
       
        if (!File.Exists(templateImagePath))
        {
            throw new FileNotFoundException("Template image was not found.", templateImagePath);
        }

        var templateMat = VisionTemplateResourceManager.Acquire(templateImagePath);
       
        try
        {
            return await FindImageAsync(templateMat, minConfidence, searchRegion, attempts, retryDelayMs, scaleStep, maxScale, minScale, cancellationToken);
        }
        finally
        {
            VisionTemplateResourceManager.Release(templateImagePath);
        }
    }

    


    /// <summary>
    /// Finds the best match for an in-memory bitmap template on screen and returns its confidence score.
    /// </summary>
    /// <param name="templateBitmap">Template bitmap to locate on screen.</param>
    /// <param name="minConfidence">
    /// Minimum accepted confidence in the range [0..1]. Higher values are stricter.
    /// </param>
    /// <param name="searchRegion">
    /// Optional region to search. When null, the full virtual desktop is searched.
    /// </param>
    /// <returns>
    /// The best image match if confidence is at least <paramref name="minConfidence"/>; otherwise null.
    /// </returns>
    /// 
    /// <exception cref="ObjectDisposedException">Thrown when this instance is disposed.</exception>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="templateBitmap"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when inputs are invalid.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when confidence is outside [0..1].</exception>
    /// <exception cref="InvalidOperationException">Thrown when template conversion fails.</exception>
    public async Task<ImageMatchResult?> FindImageAsync
    (
        Bitmap templateBitmap,
        double minConfidence = 0.8, 
        Rectangle? searchRegion = null, 
        int attempts = 10,
        int retryDelayMs = 250,
        float scaleStep = 0.05f,
        float maxScale = 2.0f,
        float minScale = 0.25f,
        CancellationToken cancellationToken = default
    )
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(templateBitmap);

        using var templateMat = ConvertBitmapToMat(templateBitmap);
        return await FindImageAsync(templateMat, minConfidence, searchRegion, attempts, retryDelayMs, scaleStep, maxScale, minScale, cancellationToken);
    }

    public async Task<ImageMatchResult?> FindImageAsync
    (
        Mat templateMat,
        double minConfidence = 0.8, 
        Rectangle? searchRegion = null, 
        int attempts = 10,
        int retryDelayMs = 250,
        float scaleStep = 0.05f,
        float maxScale = 2.0f,
        float minScale = 0.25f,
        CancellationToken cancellationToken = default
    )
    {
        ThrowIfDisposed();

        if (minConfidence is < 0 or > 1)
        {
            throw new ArgumentOutOfRangeException(nameof(minConfidence), "minConfidence must be between 0 and 1.");
        }

        if (attempts <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(attempts), "attempts must be greater than zero.");
        }

        if (retryDelayMs < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(retryDelayMs), "retryDelayMs cannot be negative.");
        }

        var region = NormalizeRegion(searchRegion);

        if (templateMat.Empty())
        {
            throw new InvalidOperationException("Template image is empty or could not be decoded.");
        }

        // Generate scales in breadth-first order around 1.0:
        // 1.0, 0.95, 1.05, 0.90, 1.10, ...
        // This prioritizes likely scale matches before exploring farther values.
        var scales = GenerateScalesInBFSOrder(minScale, maxScale, scaleStep);

        for (var attempt = 1; attempt <= attempts; attempt++)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                break;
            }

            // Capture once per attempt, then test all scales against the same frame.
            using var screenshot = CaptureScreenshot(region);
            using var searchMat = ConvertBitmapToMat(screenshot);

            // Try each scale in BFS order around 1.0.
            foreach (var scale in scales)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    break;
                }

                // Resize template to current test scale.
                Mat scaledTemplate;
                if (Math.Abs(scale - 1.0) > 0.001)
                {
                    scaledTemplate = new Mat();
                    int scaledWidth = Math.Max(1, (int)(templateMat.Width * scale));
                    int scaledHeight = Math.Max(1, (int)(templateMat.Height * scale));
                    Cv2.Resize(templateMat, scaledTemplate, new OpenCvSharp.Size(scaledWidth, scaledHeight), interpolation: InterpolationFlags.Linear);
                }
                else
                {
                    scaledTemplate = templateMat;
                }

                try
                {
                    if (scaledTemplate.Width > searchMat.Width || scaledTemplate.Height > searchMat.Height)
                    {
                        // Template cannot fit inside current search image at this scale.
                        continue;
                    }

                    var resultWidth = searchMat.Width - scaledTemplate.Width + 1;
                    var resultHeight = searchMat.Height - scaledTemplate.Height + 1;

                    // MatchTemplate produces a score map: one score for each possible template position.
                    using var resultMat = new Mat(resultHeight, resultWidth, MatType.CV_32FC1);
                    Cv2.MatchTemplate(searchMat, scaledTemplate, resultMat, _options.TemplateMatchMode);
                    Cv2.MinMaxLoc(resultMat, out var minValue, out var maxValue, out var minLocation, out var maxLocation);

                    // For SqDiff modes, lower is better so we invert to a confidence-like score.
                    var (confidence, location) = IsLowerScoreBetter(_options.TemplateMatchMode)
                        ? (1 - minValue, minLocation)
                        : (maxValue, maxLocation);

                    if (confidence >= minConfidence)
                    {
                        // Return the first confidence-qualified hit. Because scales are BFS-ordered,
                        // this tends to find near-1.0 matches first and keeps runtime predictable.
                        var localBounds = new Rectangle(
                            location.X,
                            location.Y,
                            scaledTemplate.Width,
                            scaledTemplate.Height);

                        return new ImageMatchResult(localBounds, confidence, region);
                    }
                }
                finally
                {
                    if (Math.Abs(scale - 1.0) > 0.001)
                    {
                        scaledTemplate?.Dispose();
                    }
                }
            }

            if (attempt < attempts && retryDelayMs > 0)
            {
                await Task.Delay(retryDelayMs, cancellationToken).ConfigureAwait(false);
            }
        }

        return null;
    }

    private static double[] GenerateScalesInBFSOrder(double minScale = 0.2, double maxScale = 2.0, double interval = 0.2)
    {
        if (interval <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(interval), "interval must be greater than zero.");
        }

        if (minScale <= 0 || maxScale <= 0 || minScale > maxScale)
        {
            throw new ArgumentOutOfRangeException(nameof(minScale), "Scale bounds must be positive and minScale <= maxScale.");
        }

        var scales = new List<double> { 1.0 };

        // Expand outward from 1.0 equally in both directions.
        // Example with interval 0.2: 1.0, 0.8, 1.2, 0.6, 1.4, ...
        for (double offset = interval; offset <= 1.0; offset += interval)
        {
            double below = 1.0 - offset;
            double above = 1.0 + offset;

            if (below >= minScale)
            {
                scales.Add(below);
            }

            if (above <= maxScale)
            {
                scales.Add(above);
            }
        }

        return scales.ToArray();
    }

    /// <summary>
    /// Finds a template using edge pyramid template matching, which is more robust to color/theme changes.
    /// </summary>
    /// <param name="templateImage">Template file name under the configured template folder.</param>
    /// <param name="minConfidence">Minimum accepted match confidence [0..1].</param>
    /// <param name="searchRegion">Optional area to search inside. Null means full virtual desktop.</param>
    /// <param name="attempts">How many capture+match retries to perform.</param>
    /// <param name="retryDelayMs">Delay between retries in milliseconds.</param>
    /// <param name="pyramidLevels">Number of image pyramid levels to compare.</param>
    /// <param name="cannyThreshold1">First Canny threshold for edge extraction.</param>
    /// <param name="cannyThreshold2">Second Canny threshold for edge extraction.</param>
    /// <param name="minColorCorrelation">Minimum HSV histogram correlation [0..1] for color verification.</param>
    /// <param name="cancellationToken">Cancellation token for cooperative cancel.</param>
    /// <returns>The best edge-based match if confidence qualifies; otherwise null.</returns>
    public async Task<ImageMatchResult?> FindImageByEdgePyramidAsync(
        string templateImage,
        double minConfidence = 0.8,
        Rectangle? searchRegion = null,
        int attempts = 10,
        int retryDelayMs = 250,
        int pyramidLevels = 4,
        double cannyThreshold1 = 50,
        double cannyThreshold2 = 150,
        double minColorCorrelation = 0,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrWhiteSpace(templateImage);

        var templateImagePath = Path.Combine(_options.TemplatePath, templateImage);
        if (!File.Exists(templateImagePath))
        {
            throw new FileNotFoundException("Template image was not found.", templateImagePath);
        }

        var templateMat = VisionTemplateResourceManager.Acquire(templateImagePath);
        try
        {
            return await FindImageByEdgePyramidAsync(
                templateMat,
                minConfidence,
                searchRegion,
                attempts,
                retryDelayMs,
                pyramidLevels,
                cannyThreshold1,
                cannyThreshold2,
                minColorCorrelation,
                cancellationToken);
        }
        finally
        {
            VisionTemplateResourceManager.Release(templateImagePath);
        }
    }

    /// <summary>
    /// Finds a bitmap template using edge pyramid template matching.
    /// </summary>
    /// <param name="templateBitmap">Template image in memory.</param>
    /// <param name="minConfidence">Minimum accepted match confidence [0..1].</param>
    /// <param name="searchRegion">Optional area to search inside. Null means full virtual desktop.</param>
    /// <param name="attempts">How many capture+match retries to perform.</param>
    /// <param name="retryDelayMs">Delay between retries in milliseconds.</param>
    /// <param name="pyramidLevels">Number of image pyramid levels to compare.</param>
    /// <param name="cannyThreshold1">First Canny threshold for edge extraction.</param>
    /// <param name="cannyThreshold2">Second Canny threshold for edge extraction.</param>
    /// <param name="minColorCorrelation">Minimum HSV histogram correlation [0..1] for color verification.</param>
    /// <param name="cancellationToken">Cancellation token for cooperative cancel.</param>
    /// <returns>The best edge-based match if confidence qualifies; otherwise null.</returns>
    public async Task<ImageMatchResult?> FindImageByEdgePyramidAsync(
        Bitmap templateBitmap,
        double minConfidence = 0.8,
        Rectangle? searchRegion = null,
        int attempts = 10,
        int retryDelayMs = 250,
        int pyramidLevels = 4,
        double cannyThreshold1 = 50,
        double cannyThreshold2 = 150,
        double minColorCorrelation = 0.0,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(templateBitmap);

        using var templateMat = ConvertBitmapToMat(templateBitmap);
        return await FindImageByEdgePyramidAsync(
            templateMat,
            minConfidence,
            searchRegion,
            attempts,
            retryDelayMs,
            pyramidLevels,
            cannyThreshold1,
            cannyThreshold2,
            minColorCorrelation,
            cancellationToken);
    }

    /// <summary>
    /// Finds an in-memory template using edge pyramid template matching.
    /// </summary>
    /// <param name="templateMat">Template image in OpenCV matrix form.</param>
    /// <param name="minConfidence">Minimum accepted match confidence [0..1].</param>
    /// <param name="searchRegion">Optional area to search inside. Null means full virtual desktop.</param>
    /// <param name="attempts">How many capture+match retries to perform.</param>
    /// <param name="retryDelayMs">Delay between retries in milliseconds.</param>
    /// <param name="pyramidLevels">Number of image pyramid levels to compare.</param>
    /// <param name="cannyThreshold1">First Canny threshold for edge extraction.</param>
    /// <param name="cannyThreshold2">Second Canny threshold for edge extraction.</param>
    /// <param name="minColorCorrelation">Minimum HSV histogram correlation [0..1] for color verification.</param>
    /// <param name="cancellationToken">Cancellation token for cooperative cancel.</param>
    /// <returns>The best edge-based match if confidence qualifies; otherwise null.</returns>
    public async Task<ImageMatchResult?> FindImageByEdgePyramidAsync(
        Mat templateMat,
        double minConfidence = 0.8,
        Rectangle? searchRegion = null,
        int attempts = 10,
        int retryDelayMs = 250,
        int pyramidLevels = 4,
        double cannyThreshold1 = 50,
        double cannyThreshold2 = 150,
        double minColorCorrelation = 0.0,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();

        if (minConfidence is < 0 or > 1)
        {
            throw new ArgumentOutOfRangeException(nameof(minConfidence), "minConfidence must be between 0 and 1.");
        }

        if (attempts <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(attempts), "attempts must be greater than zero.");
        }

        if (retryDelayMs < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(retryDelayMs), "retryDelayMs cannot be negative.");
        }

        if (pyramidLevels <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(pyramidLevels), "pyramidLevels must be greater than zero.");
        }

        if (minColorCorrelation is < 0 or > 1)
        {
            throw new ArgumentOutOfRangeException(nameof(minColorCorrelation), "minColorCorrelation must be between 0 and 1.");
        }

        if (templateMat.Empty())
        {
            throw new InvalidOperationException("Template image is empty or could not be decoded.");
        }

        var region = NormalizeRegion(searchRegion);

        // Build template edges once, then construct template pyramid once.
        // This is reused across attempts to avoid repeated preprocessing.
        using var templateEdges = BuildEdgeMap(templateMat, cannyThreshold1, cannyThreshold2);
        var templatePyramid = BuildPyramid(templateEdges, pyramidLevels);

        try
        {
            for (var attempt = 1; attempt <= attempts; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                using var screenshot = CaptureScreenshot(region);
                using var searchMat = ConvertBitmapToMat(screenshot);
                using var searchEdges = BuildEdgeMap(searchMat, cannyThreshold1, cannyThreshold2);

                // Rebuild search pyramid every attempt because screen content may change.
                var searchPyramid = BuildPyramid(searchEdges, pyramidLevels);
                try
                {
                    double bestConfidence = -1;
                    Rectangle bestBounds = Rectangle.Empty;

                    for (var level = 0; level < pyramidLevels; level++)
                    {
                        var templateLevel = templatePyramid[level];
                        var searchLevel = searchPyramid[level];

                        if (templateLevel.Width > searchLevel.Width || templateLevel.Height > searchLevel.Height)
                        {
                            continue;
                        }

                        var resultWidth = searchLevel.Width - templateLevel.Width + 1;
                        var resultHeight = searchLevel.Height - templateLevel.Height + 1;

                        using var resultMat = new Mat(resultHeight, resultWidth, MatType.CV_32FC1);
                        Cv2.MatchTemplate(searchLevel, templateLevel, resultMat, _options.TemplateMatchMode);
                        Cv2.MinMaxLoc(resultMat, out var minValue, out var maxValue, out var minLocation, out var maxLocation);

                        var (confidence, location) = IsLowerScoreBetter(_options.TemplateMatchMode)
                            ? (1 - minValue, minLocation)
                            : (maxValue, maxLocation);

                        if (confidence <= bestConfidence)
                        {
                            continue;
                        }

                        // Level N in a pyramid is downsampled by 2^N in each axis,
                        // so convert local coordinates back to level-0 coordinates.
                        var scaleFactor = 1 << level;
                        bestConfidence = confidence;
                        bestBounds = new Rectangle(
                            location.X * scaleFactor,
                            location.Y * scaleFactor,
                            templateLevel.Width * scaleFactor,
                            templateLevel.Height * scaleFactor);
                    }

                    if (bestConfidence >= minConfidence && bestBounds.Width > 0 && bestBounds.Height > 0)
                    {
                        // Edge matching is shape-focused; this color verification step rejects
                        // similarly shaped false positives that differ in color distribution.
                        var clampedBounds = ClampToImageBounds(bestBounds, searchMat.Width, searchMat.Height);
                        if (clampedBounds == Rectangle.Empty)
                        {
                            continue;
                        }

                        // Color verification is optional. Use a positive threshold to enable it.
                        if (minColorCorrelation > 0)
                        {
                            var colorCorrelation = ComputeHsvColorCorrelation(templateMat, searchMat, clampedBounds);
                            if (colorCorrelation < minColorCorrelation)
                            {
                                continue;
                            }
                        }

                        return new ImageMatchResult(clampedBounds, bestConfidence, region);
                    }
                }
                finally
                {
                    // Each pyramid level owns unmanaged buffers; release promptly.
                    foreach (var mat in searchPyramid)
                    {
                        mat.Dispose();
                    }
                }

                if (attempt < attempts && retryDelayMs > 0)
                {
                    await Task.Delay(retryDelayMs, cancellationToken).ConfigureAwait(false);
                }
            }
        }
        finally
        {
            // Template pyramid is shared across attempts and disposed once here.
            foreach (var mat in templatePyramid)
            {
                mat.Dispose();
            }
        }

        return null;
    }

    /// <summary>
    /// Reads all visible text from screen (or a region) and returns OCR words with confidence.
    /// </summary>
    /// <param name="searchRegion">
    /// Optional region to OCR. When null, the full virtual desktop is processed.
    /// </param>
    /// <returns>The aggregated OCR text plus per-word text, bounds, and confidence.</returns>
    /// <exception cref="ObjectDisposedException">Thrown when this instance is disposed.</exception>
    public OcrResult ReadText(Rectangle? searchRegion = null)
    {
        ThrowIfDisposed();

        var region = NormalizeRegion(searchRegion);
        // Capture pixels, run OCR, then project OCR boxes back into screen coordinates.
        using var screenshot = CaptureScreenshot(region);
        using var page = ProcessOcr(screenshot);
        var fullText = page.GetText() ?? string.Empty;
        var words = ExtractWords(page, region);

        return new OcrResult(fullText, words, region);
    }

    /// <summary>
    /// Finds OCR words that contain the provided text and meet the minimum confidence.
    /// </summary>
    /// <param name="text">Text fragment to find in OCR words.</param>
    /// <param name="minConfidence">Minimum OCR confidence in range [0..1].</param>
    /// <param name="searchRegion">
    /// Optional region to OCR. When null, the full virtual desktop is processed.
    /// </param>
    /// <param name="comparison">String comparison behavior used by <see cref="string.Contains(string, StringComparison)"/>.</param>
    /// <returns>All OCR word matches that satisfy text and confidence filters.</returns>
    /// <exception cref="ObjectDisposedException">Thrown when this instance is disposed.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="text"/> is invalid.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when confidence is outside [0..1].</exception>
    public async Task<IReadOnlyList<TextMatchResult>> FindTextAsync
    (
        string text,
        double minConfidence = 0.5,
        Rectangle? searchRegion = null,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase,
        int attempts = 10,
        int retryDelayMs = 250,
        CancellationToken cancellationToken = default
    )
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrWhiteSpace(text);

        if (minConfidence is < 0 or > 1)
        {
            throw new ArgumentOutOfRangeException(nameof(minConfidence), "minConfidence must be between 0 and 1.");
        }

        if (attempts <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(attempts), "attempts must be greater than zero.");
        }

        if (retryDelayMs < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(retryDelayMs), "retryDelayMs cannot be negative.");
        }

        for (var attempt = 1; attempt <= attempts; attempt++)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                break;
            }

            var ocrResult = ReadText(searchRegion);
            var matches = ocrResult.Words
                .Where(word => word.Confidence >= minConfidence && word.Text.Contains(text, comparison))
                .Select(word => new TextMatchResult(word.Text, word.Bounds, word.Confidence, word.SearchRegion))
                .ToList();

            if (matches.Count > 0)
            {
                return matches;
            }

            if (attempt < attempts && retryDelayMs > 0)
            {
                await Task.Delay(retryDelayMs, cancellationToken).ConfigureAwait(false);
            }
        }

        return Array.Empty<TextMatchResult>();
    }

    private IReadOnlyList<OcrWord> ExtractWords(Page page, Rectangle region)
    {
        var words = new List<OcrWord>();

        // Iterate OCR results at word level to keep location and confidence metadata.
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

            // Tesseract exposes confidence as percentage [0..100], normalize to [0..1].
            var confidence = iterator.GetConfidence(PageIteratorLevel.Word) / 100.0;
            var localBounds = new Rectangle(
                box.X1,
                box.Y1,
                box.Width,
                box.Height);

            words.Add(new OcrWord(rawText.Trim(), localBounds, confidence, region));
        }
        while (iterator.Next(PageIteratorLevel.Word));

        return words;
    }

    private Page ProcessOcr(Bitmap screenshot)
    {
        // Tesseract engine is not thread-safe; serialize access.
        lock (_ocrLock)
        {
            EnsureOcrEngine();

            using var pix = ConvertBitmapToPix(screenshot);
            return _ocrEngine!.Process(pix);
        }
    }

    private static Mat ConvertBitmapToMat(Bitmap bitmap)
    {
        // Encode to PNG bytes, then decode through OpenCV to avoid GDI/OpenCV interop issues.
        using var memoryStream = new MemoryStream();
        bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
        var imageBytes = memoryStream.ToArray();
        return Cv2.ImDecode(imageBytes, ImreadModes.Color);
    }

    private static Pix ConvertBitmapToPix(Bitmap bitmap)
    {
        // Encode to bytes and let Tesseract load a Pix from in-memory data.
        using var memoryStream = new MemoryStream();
        bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
        var imageBytes = memoryStream.ToArray();
        return Pix.LoadFromMemory(imageBytes);
    }

    private void EnsureOcrEngine()
    {
        // Lazily initialize the OCR engine to avoid startup overhead until OCR is requested.
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
        // Validate explicit region or fall back to virtual desktop bounds.
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
            throw new InvalidOperationException("Could not determine virtual desktop screen bounds.");
        }

        return new Rectangle(left, top, width, height);
    }

    private static bool IsLowerScoreBetter(TemplateMatchModes mode)
    {
        // SqDiff modes represent better matches with lower values.
        return mode is TemplateMatchModes.SqDiff or TemplateMatchModes.SqDiffNormed;
    }

    private static Mat BuildEdgeMap(Mat source, double cannyThreshold1, double cannyThreshold2)
    {
        // Normalize to grayscale first to keep edge extraction consistent across color themes.
        using var gray = new Mat();
        if (source.Channels() == 1)
        {
            source.CopyTo(gray);
        }
        else
        {
            Cv2.CvtColor(source, gray, ColorConversionCodes.BGR2GRAY);
        }

        using var blurred = new Mat();
        // Light blur reduces tiny texture noise that can destabilize Canny edges.
        Cv2.GaussianBlur(gray, blurred, new OpenCvSharp.Size(3, 3), 0);

        var edges = new Mat();
        Cv2.Canny(blurred, edges, cannyThreshold1, cannyThreshold2);
        return edges;
    }

    /// <summary>
    /// Builds an image pyramid where each next level is downsampled by 2.
    /// </summary>
    /// <param name="source">Input image.</param>
    /// <param name="levels">Maximum number of levels to create.</param>
    /// <returns>List of level images from full size to smallest level.</returns>
    private static List<Mat> BuildPyramid(Mat source, int levels)
    {
        var pyramid = new List<Mat>(levels);
        var level0 = source.Clone();
        pyramid.Add(level0);

        for (var level = 1; level < levels; level++)
        {
            if (pyramid[level - 1].Width <= 1 || pyramid[level - 1].Height <= 1)
            {
                break;
            }

            var next = new Mat();
            Cv2.PyrDown(pyramid[level - 1], next);
            pyramid.Add(next);
        }

        return pyramid;
    }

    private static Rectangle ClampToImageBounds(Rectangle bounds, int imageWidth, int imageHeight)
    {
        var left = Math.Clamp(bounds.Left, 0, imageWidth);
        var top = Math.Clamp(bounds.Top, 0, imageHeight);
        var right = Math.Clamp(bounds.Right, 0, imageWidth);
        var bottom = Math.Clamp(bounds.Bottom, 0, imageHeight);

        var width = right - left;
        var height = bottom - top;
        if (width <= 0 || height <= 0)
        {
            return Rectangle.Empty;
        }

        return new Rectangle(left, top, width, height);
    }

    private static double ComputeHsvColorCorrelation(Mat templateMat, Mat searchMat, Rectangle candidateBounds)
    {
        using var searchRoi = new Mat(searchMat, new OpenCvSharp.Rect(candidateBounds.X, candidateBounds.Y, candidateBounds.Width, candidateBounds.Height));
        using var resizedTemplate = new Mat();
        Cv2.Resize(templateMat, resizedTemplate, new OpenCvSharp.Size(candidateBounds.Width, candidateBounds.Height), interpolation: InterpolationFlags.Linear);

        using var templateBgr = new Mat();
        if (resizedTemplate.Channels() == 1)
        {
            Cv2.CvtColor(resizedTemplate, templateBgr, ColorConversionCodes.GRAY2BGR);
        }
        else
        {
            resizedTemplate.CopyTo(templateBgr);
        }

        using var searchBgr = new Mat();
        if (searchRoi.Channels() == 1)
        {
            Cv2.CvtColor(searchRoi, searchBgr, ColorConversionCodes.GRAY2BGR);
        }
        else
        {
            searchRoi.CopyTo(searchBgr);
        }

        using var templateHsv = new Mat();
        using var searchHsv = new Mat();
        Cv2.CvtColor(templateBgr, templateHsv, ColorConversionCodes.BGR2HSV);
        Cv2.CvtColor(searchBgr, searchHsv, ColorConversionCodes.BGR2HSV);

        // Low-saturation UI (white/gray text/buttons) makes hue unstable.
        // Use Value-channel histograms in that case; otherwise blend HS + V.
        var templateMean = Cv2.Mean(templateHsv);
        var searchMean = Cv2.Mean(searchHsv);
        var lowSaturation = templateMean.Val1 < 25 && searchMean.Val1 < 25;

        var valueCorrelation = ComputeHistogramCorrelation(
            templateHsv,
            searchHsv,
            [2],
            [32],
            [new Rangef(0, 256)]);

        double score;
        if (lowSaturation)
        {
            score = valueCorrelation;
        }
        else
        {
            var hsCorrelation = ComputeHistogramCorrelation(
                templateHsv,
                searchHsv,
                [0, 1],
                [30, 32],
                [new Rangef(0, 180), new Rangef(0, 256)]);
            score = (0.75 * hsCorrelation) + (0.25 * valueCorrelation);
        }

        if (double.IsNaN(score) || double.IsInfinity(score))
        {
            return -1;
        }

        return score;
    }

    private static double ComputeHistogramCorrelation(
        Mat templateHsv,
        Mat searchHsv,
        int[] channels,
        int[] histSize,
        Rangef[] ranges)
    {
        using var templateHist = new Mat();
        using var searchHist = new Mat();
        Cv2.CalcHist([templateHsv], channels, null, templateHist, channels.Length, histSize, ranges);
        Cv2.CalcHist([searchHsv], channels, null, searchHist, channels.Length, histSize, ranges);

        if (Cv2.CountNonZero(templateHist) == 0 || Cv2.CountNonZero(searchHist) == 0)
        {
            return -1;
        }

        Cv2.Normalize(templateHist, templateHist, 1, 0, NormTypes.L1);
        Cv2.Normalize(searchHist, searchHist, 1, 0, NormTypes.L1);

        return Cv2.CompareHist(templateHist, searchHist, HistCompMethods.Correl);
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

        // Release unmanaged OCR resources.
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

/// <summary>
/// Provides helper conversions between region-local coordinates and global capture-space coordinates.
/// </summary>
public static class VisionCoordinateSpace
{
    /// <summary>
    /// Converts local bounds (relative to a search region) into global bounds.
    /// </summary>
    public static Rectangle ToGlobalBounds(Rectangle localBounds, Rectangle searchRegion)
    {
        return new Rectangle(
            searchRegion.Left + localBounds.Left,
            searchRegion.Top + localBounds.Top,
            localBounds.Width,
            localBounds.Height);
    }

    /// <summary>
    /// Converts global bounds into local bounds relative to a search region.
    /// </summary>
    public static Rectangle ToLocalBounds(Rectangle globalBounds, Rectangle searchRegion)
    {
        return new Rectangle(
            globalBounds.Left - searchRegion.Left,
            globalBounds.Top - searchRegion.Top,
            globalBounds.Width,
            globalBounds.Height);
    }
}

/// <summary>
/// Represents an image template match on screen.
/// </summary>
/// <param name="Bounds">Matched region bounds local to the <paramref name="SearchRegion"/>.</param>
/// <param name="Confidence">Match confidence in the range [0..1].</param>
/// <param name="SearchRegion">Search region used during detection, in capture-space coordinates.</param>
public sealed record ImageMatchResult(
    Rectangle Bounds,
    double Confidence,
    Rectangle SearchRegion = default)
{
    /// <summary>
    /// Converts this local match bounds to global virtual-desktop coordinates.
    /// </summary>
    public Rectangle ToGlobalBounds()
    {
        return VisionCoordinateSpace.ToGlobalBounds(Bounds, SearchRegion);
    }
}

/// <summary>
/// Represents one OCR word with location and confidence.
/// </summary>
/// <param name="Text">Recognized word text.</param>
/// <param name="Bounds">Word bounds local to the <paramref name="SearchRegion"/>.</param>
/// <param name="Confidence">OCR confidence in the range [0..1].</param>
/// <param name="SearchRegion">OCR region used during detection, in global virtual-desktop coordinates.</param>
public sealed record OcrWord(
    string Text,
    Rectangle Bounds,
    double Confidence,
    Rectangle SearchRegion = default)
{
    /// <summary>
    /// Converts this local word bounds to global virtual-desktop coordinates.
    /// </summary>
    public Rectangle ToGlobalBounds()
    {
        return VisionCoordinateSpace.ToGlobalBounds(Bounds, SearchRegion);
    }
}

/// <summary>
/// Represents OCR output for a processed region.
/// </summary>
/// <param name="Text">Full OCR text content for the region.</param>
/// <param name="Words">Word-level OCR results with bounds and confidence.</param>
/// <param name="SearchRegion">OCR region used during detection, in global virtual-desktop coordinates.</param>
public sealed record OcrResult(
    string Text,
    IReadOnlyList<OcrWord> Words,
    Rectangle SearchRegion = default);

/// <summary>
/// Represents a text match derived from OCR words.
/// </summary>
/// <param name="Text">Matched text segment.</param>
/// <param name="Bounds">Matched bounds local to the <paramref name="SearchRegion"/>.</param>
/// <param name="Confidence">OCR confidence in the range [0..1].</param>
/// <param name="SearchRegion">OCR region used during detection, in global virtual-desktop coordinates.</param>
public sealed record TextMatchResult(
    string Text,
    Rectangle Bounds,
    double Confidence,
    Rectangle SearchRegion = default)
{
    /// <summary>
    /// Converts this local text-match bounds to global virtual-desktop coordinates.
    /// </summary>
    public Rectangle ToGlobalBounds()
    {
        return VisionCoordinateSpace.ToGlobalBounds(Bounds, SearchRegion);
    }
}