
using System.Diagnostics;
using AutomationFramework.Windows;
using AutomationRunner.Scripting;
using Microsoft.Extensions.Logging;


namespace AutomationRunner.Scripts;

/// <summary>
/// Requires a wow shortcut on the background called "WoW"
/// </summary>
public sealed class StartWoW : BaseScript
{
    public StartWoW(ILogger<StartWoW> logger)
        : base(logger)
    {
    }

    public override string Name => "start-wow";

    public override string Description => "Starts the WoW Application";


    protected override Task InitializeAsync(CancellationToken cancellationToken)
    {
        return Task.CompletedTask;
    }
    
    public override void Dispose()
    {

    }

    protected override async Task RunAsync(CancellationToken cancellationToken)
    {
        string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string shortcutPath = Path.Combine(desktop, "WoW.lnk");
        
        Logger.LogInformation("Starting WoW from shortcut: {ShortcutPath}", shortcutPath);

        var wowProcess = Process.Start(new ProcessStartInfo
        {
            FileName = shortcutPath,
            UseShellExecute = true
        });

        if (wowProcess is null)
        {
            throw new InvalidOperationException($"Failed to start WoW process from shortcut {shortcutPath}.");
        }

        //give it some time to start up
        await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);


        if (WinWindowManager.GetForegroundWindowHandle() != wowProcess.Handle)
        {
            WinWindowManager.FocusWindow(wowProcess.Handle);
            Logger.LogInformation("WoW window focused successfully.");
        }
    }
}

