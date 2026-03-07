using AutomationRunner.Scripting;
using Microsoft.Extensions.Configuration;
using System.CommandLine;

namespace AutomationRunner;

internal static class Program
{
    private static int Main(string[] args)
    {
        var configuration = BuildConfiguration();
        var scripts = ScriptRegistry.DiscoverScripts();

        if (scripts.Count == 0)
        {
            Console.WriteLine("No scripts were found. Add classes in the Scripts folder that implement IAutomationScript.");
            return 1;
        }

        var rootCommand = new RootCommand("AutomationRunner script runner")
        {
            CreateRunCommand(configuration, scripts),
            CreateListCommand(configuration, scripts)
        };

        return rootCommand.Parse(args).Invoke();
    }

    private static Command CreateRunCommand(IConfiguration configuration, IReadOnlyList<IAutomationScript> scripts)
    {
        var scriptNameArgument = new Argument<string>("script-name")
        {
            Description = "The script name to run."
        };

        var runCommand = new Command("run", "Run a script by name.")
        {
            scriptNameArgument
        };
        
        runCommand.SetAction(parseResult =>
        {
            var requestedScriptName = parseResult.GetValue(scriptNameArgument);
            return HandleRunCommandAsync(requestedScriptName, configuration, scripts, CancellationToken.None)
                .GetAwaiter()
                .GetResult();
        });

        return runCommand;
    }

    private static Command CreateListCommand(IConfiguration configuration, IReadOnlyList<IAutomationScript> scripts)
    {
        var listCommand = new Command("list", "List scripts, then choose one to run or exit.");
        listCommand.SetAction(_ =>
            HandleListCommandAsync(configuration, scripts, CancellationToken.None)
                .GetAwaiter()
                .GetResult());

        return listCommand;
    }

    private static async Task<int> HandleRunCommandAsync(
        string scriptName,
        IConfiguration configuration,
        IReadOnlyList<IAutomationScript> scripts,
        CancellationToken cancellationToken)
    {
        var selectedScript = ResolveByName(scriptName, scripts);
        if (selectedScript is null)
        {
            Console.WriteLine($"Script '{scriptName}' was not found.");
            return 1;
        }

        return await RunScriptAsync(selectedScript, configuration, cancellationToken);
    }

    private static async Task<int> HandleListCommandAsync(
        IConfiguration configuration,
        IReadOnlyList<IAutomationScript> scripts,
        CancellationToken cancellationToken)
    {
        while (true)
        {
            var selection = PromptForScriptSelection(scripts);
            if (selection.ExitRequested)
            {
                Console.WriteLine("Exiting.");
                return 0;
            }

            if (selection.Script is null)
            {
                Console.WriteLine("No script selected.");
                Console.WriteLine();
                continue;
            }

            return await RunScriptAsync(selection.Script, configuration, cancellationToken);
        }
    }

    private static async Task<int> RunScriptAsync(
        IAutomationScript selectedScript,
        IConfiguration configuration,
        CancellationToken cancellationToken)
    {
        using var cancellationTokenSource = new CancellationTokenSource();
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationTokenSource.Token, cancellationToken);

        void OnCancelKeyPress(object? _, ConsoleCancelEventArgs eventArgs)
        {
            eventArgs.Cancel = true;
            cancellationTokenSource.Cancel();
        }

        Console.CancelKeyPress += OnCancelKeyPress;

        var context = new ScriptExecutionContext(configuration);

        Console.WriteLine($"Running script: {selectedScript.Name}");
        Console.WriteLine("Press Ctrl+C to stop.");

        try
        {
            await selectedScript.ExecuteAsync(context, linkedCts.Token);
            Console.WriteLine("Script completed.");
            Console.WriteLine();
            return 0;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Script canceled.");
            return 2;
        }
        finally
        {
            Console.CancelKeyPress -= OnCancelKeyPress;
        }
    }

    private static IConfiguration BuildConfiguration()
    {
        var environment = Environment.GetEnvironmentVariable("DOTNET_ENVIRONMENT")
            ?? Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT")
            ?? "Production";

        return new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddJsonFile($"appsettings.{environment}.json", optional: true, reloadOnChange: false)
            .Build();
    }

    private static ScriptSelection PromptForScriptSelection(IReadOnlyList<IAutomationScript> scripts)
    {
        PrintScriptList(scripts);
        Console.Write("Select a script by number or name (or type 'exit'): ");

        var input = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(input))
        {
            return new ScriptSelection(null, false);
        }

        var selector = input.Trim();
        if (string.Equals(selector, "exit", StringComparison.OrdinalIgnoreCase)
            || string.Equals(selector, "quit", StringComparison.OrdinalIgnoreCase)
            || string.Equals(selector, "q", StringComparison.OrdinalIgnoreCase)
            || string.Equals(selector, "x", StringComparison.OrdinalIgnoreCase))
        {
            return new ScriptSelection(null, true);
        }

        var script = ResolveByNameOrIndex(selector, scripts);
        if (script is null)
        {
            Console.WriteLine($"Script '{input}' was not found.");
        }

        return new ScriptSelection(script, false);
    }

    private static IAutomationScript? ResolveByName(string scriptName, IReadOnlyList<IAutomationScript> scripts)
    {
        return scripts.FirstOrDefault(script =>
            string.Equals(script.Name, scriptName, StringComparison.OrdinalIgnoreCase));
    }

    private static IAutomationScript? ResolveByNameOrIndex(string selector, IReadOnlyList<IAutomationScript> scripts)
    {
        if (int.TryParse(selector, out var scriptNumber))
        {
            var index = scriptNumber - 1;
            if (index >= 0 && index < scripts.Count)
            {
                return scripts[index];
            }

            Console.WriteLine($"Script number '{scriptNumber}' is out of range.");
            return null;
        }

        return scripts.FirstOrDefault(script => string.Equals(script.Name, selector, StringComparison.OrdinalIgnoreCase));
    }

    private static void PrintScriptList(IReadOnlyList<IAutomationScript> scripts)
    {
        Console.WriteLine("Available scripts:");
        for (var index = 0; index < scripts.Count; index++)
        {
            var script = scripts[index];
            Console.WriteLine($"  {index + 1}. {script.Name} - {script.Description}");
        }
    }

    private readonly record struct ScriptSelection(IAutomationScript? Script, bool ExitRequested);

}