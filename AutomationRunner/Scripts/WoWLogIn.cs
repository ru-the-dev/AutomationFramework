using System;
using System.Reflection.Metadata;
using System.Reflection.Metadata.Ecma335;
using AutomationFramework;
using AutomationFramework.Extensions;
using AutomationRunner.Scripting;
using AutomationRunner.Services;
using Microsoft.Extensions.Logging;

namespace AutomationRunner.Scripts;

public class WoWLogIn : BaseScript
{
    public WoWLogIn
    (
        ILogger<WoWLogIn> logger, 
        IAutomationVisionFactory visionFactory,
        AutomationFramework.Cursor cursor,
        Keyboard keyboard
    ) 
    : base(logger)
    {
        _visionFactory = visionFactory;
        _cursor = cursor;
        _keyboard = keyboard;
    }

    public override string Name => "wow-login";

    public override string Description => "Log into wow with credentials stored in the appsettings.";

    private readonly Keyboard _keyboard;
    private readonly AutomationFramework.Cursor _cursor;
    private readonly IAutomationVisionFactory _visionFactory;
    private Vision? _vision = null;

    private Dictionary<string, VisionTemplateLease> _templateLeases = null!;

    public override void Dispose()
    {
        // Release templates
        if (_templateLeases != null)
        {
            foreach (var templateLease in _templateLeases.Values)
            {
                templateLease?.Dispose();
            }

            _templateLeases.Clear();
        }
        
        _vision?.Dispose();
    }

    protected override Task InitializeAsync(CancellationToken cancellationToken)
    {
        _vision = _visionFactory.Create();
        
        _templateLeases = _vision.AcquireTemplateLeases(
            VisionTemplateFileNames.MAIN_MENU_PASSWORD_TEXT,
            VisionTemplateFileNames.MAIN_MENU_BUTTONS
        );


        return Task.CompletedTask;
    }

    protected override async Task RunAsync(CancellationToken cancellationToken)
    {
        if (_vision == null)
        {
            throw new InvalidOperationException("Vision system is not initialized.");
        }

        if (_templateLeases == null)
        {
            throw new InvalidOperationException("Template leases are not initialized.");
        }
        
        bool foundMainMenuButtons = await AttemptToFindMainMenuButtons(cancellationToken);
        if (foundMainMenuButtons == false)
        {
            throw new InvalidOperationException("Failed to find the main menu buttons after multiple attempts. Is WoW running and visible?");
        }

        // find password text
        var passwordMatch = await _vision.FindImageAsync(_templateLeases[VisionTemplateFileNames.MAIN_MENU_PASSWORD_TEXT].TemplateMat, 0.6);

        if (passwordMatch is null)
        {
            throw new InvalidOperationException("Failed to find the password text on the main menu. Cannot proceed with login.");
        }

        // move cursor to password box and highlight to enter password
        var passwordBoxPos = passwordMatch.ToGlobalBounds().Translate(0, (int)(passwordMatch.Bounds.Height * 1.5f)).Center();

        await _cursor.MoveToAsync(passwordBoxPos, cancellationToken: cancellationToken);
        await Task.Delay(500);

        await _cursor.ClickAsync();

        await _keyboard.TypeTextAsync("Password");
        await _keyboard.PressKeyAsync(VirtualKey.Enter);


    }


    private async Task<bool> AttemptToFindMainMenuButtons(CancellationToken cancellationToken, int maxAttempts = 3 )
    {
        ArgumentNullException.ThrowIfNull(_vision);

        for (var attempt = 1; attempt <= maxAttempts; ++attempt)
        {
            var mainMenuButtonsMatch = await _vision.FindImageAsync(
                _templateLeases[VisionTemplateFileNames.MAIN_MENU_BUTTONS].TemplateMat,
                0.6,
                cancellationToken: cancellationToken);

            if (mainMenuButtonsMatch != null)
            {
                return true;
            }

            if (attempt < maxAttempts)
            {
                await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
            }
        }

        return false;
    }
}
