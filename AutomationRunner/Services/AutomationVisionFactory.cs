using AutomationFramework;
using Microsoft.Extensions.Configuration;
using OpenCvSharp;

namespace AutomationRunner.Services;

public sealed class AutomationVisionFactory : IAutomationVisionFactory
{
    private readonly string _ocrDataPath;
    private readonly IVisionTemplateResourceManager _templateResourceManager;

    public AutomationVisionFactory(IConfiguration configuration, IVisionTemplateResourceManager templateResourceManager)
    {
        _ocrDataPath = configuration.GetRequiredSection("OcrDataPath").Value
            ?? throw new InvalidOperationException("OcrDataPath must be configured.");
        _templateResourceManager = templateResourceManager;
    }

    public Vision Create(TemplateMatchModes templateMatchMode = TemplateMatchModes.CCoeffNormed)
    {
        return new Vision(
            _templateResourceManager,
            new Vision.Options
            {
                OcrLanguage = "eng",
                OcrDataPath = _ocrDataPath,
                TemplateMatchMode = templateMatchMode
            });
    }
}
