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

        var listOption = new Option<bool>("--list", "-l") 
        {
            Description = "List discovered scripts and exit."
        };

        var scriptOption = new Option<string?>("--script", "-s")
        {
            Description = "The name or number of the script to run. If not specified, a selection prompt will be shown."
        };

        var rootCommand = new RootCommand("AutomationRunner script runner");
        rootCommand.Options.Add(listOption);
        rootCommand.Options.Add(scriptOption);

        rootCommand.SetAction(parseResult =>
        {
            var shouldList = parseResult.GetValue(listOption);
            var requestedScript = parseResult.GetValue(scriptOption);

            return ExecuteAsync(
                    shouldList,
                    requestedScript,
                    configuration,
                    scripts,
                    CancellationToken.None)
                .GetAwaiter()
                .GetResult();
        });

        return rootCommand.Parse(args).Invoke();
    }

    private static async Task<int> ExecuteAsync(
        bool shouldList,
        string? requestedScript,
        IConfiguration configuration,
        IReadOnlyList<IAutomationScript> scripts,
        CancellationToken cancellationToken)
    {
        if (shouldList)
        {
            PrintScriptList(scripts);
            return 0;
        }

        using var cancellationTokenSource = new CancellationTokenSource();
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationTokenSource.Token, cancellationToken);
        Console.CancelKeyPress += (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            cancellationTokenSource.Cancel();
        };

        var context = new ScriptExecutionContext(configuration);
        var isFirstIteration = true;

        while (!linkedCts.Token.IsCancellationRequested)
        {
            IAutomationScript? selectedScript = null;

            if (isFirstIteration)
            {
                selectedScript = ResolveScriptFromSelector(requestedScript, scripts);
            }

            if (selectedScript is null)
            {
                var selection = PromptForScriptSelection(scripts);
                if (selection.ExitRequested)
                {
                    Console.WriteLine("Exiting.");
                    return 0;
                }

                selectedScript = selection.Script;
                if (selectedScript is null)
                {
                    Console.WriteLine("No script selected.");
                    Console.WriteLine();
                    isFirstIteration = false;
                    continue;
                }
            }

            isFirstIteration = false;


            Console.WriteLine($"Running script: {selectedScript.Name}");
            Console.WriteLine("Press Ctrl+C to stop.");

            try
            {
                await selectedScript.ExecuteAsync(context, linkedCts.Token);
                Console.WriteLine("Script completed.");
                Console.WriteLine();
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Script canceled.");
                return 2;
            }
        }

        Console.WriteLine("Script canceled.");
        return 2;
    }

    private static IAutomationScript? ResolveScriptFromSelector(string? selector, IReadOnlyList<IAutomationScript> scripts)
    {
        if (string.IsNullOrWhiteSpace(selector))
        {
            return null;
        }

        return ResolveByNameOrIndex(selector, scripts);
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