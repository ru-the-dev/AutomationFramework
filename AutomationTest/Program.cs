using AutomationTest.Scripting;

namespace AutomationTest;

internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        var scripts = ScriptRegistry.DiscoverScripts();
        if (scripts.Count == 0)
        {
            Console.WriteLine("No scripts were found. Add classes in the Scripts folder that implement IAutomationScript.");
            return 1;
        }

        if (HasAnyFlag(args, "--help", "-h"))
        {
            PrintUsage();
            return 0;
        }

        if (HasAnyFlag(args, "--list", "-l"))
        {
            PrintScriptList(scripts);
            return 0;
        }

        var selectedScript = ResolveScriptFromArgs(args, scripts) ?? PromptForScriptSelection(scripts);
        if (selectedScript is null)
        {
            Console.WriteLine("No script selected.");
            return 1;
        }

        using var cancellationTokenSource = new CancellationTokenSource();
        Console.CancelKeyPress += (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            cancellationTokenSource.Cancel();
        };

        var context = new ScriptExecutionContext();

        Console.WriteLine($"Running script: {selectedScript.Name}");
        Console.WriteLine("Press Ctrl+C to stop.");

        try
        {
            await selectedScript.ExecuteAsync(context, cancellationTokenSource.Token);
            Console.WriteLine("Script completed.");
            return 0;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Script canceled.");
            return 2;
        }
    }

    private static bool HasAnyFlag(string[] args, params string[] flags)
    {
        return args.Any(argument => flags.Any(flag => string.Equals(argument, flag, StringComparison.OrdinalIgnoreCase)));
    }

    private static IAutomationScript? ResolveScriptFromArgs(IReadOnlyList<string> args, IReadOnlyList<IAutomationScript> scripts)
    {
        for (var index = 0; index < args.Count; index++)
        {
            if (!string.Equals(args[index], "--script", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(args[index], "-s", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (index + 1 >= args.Count)
            {
                Console.WriteLine("Missing value after --script.");
                return null;
            }

            return ResolveByNameOrIndex(args[index + 1], scripts);
        }

        return null;
    }

    private static IAutomationScript? PromptForScriptSelection(IReadOnlyList<IAutomationScript> scripts)
    {
        PrintScriptList(scripts);
        Console.Write("Select a script by number or name: ");

        var input = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(input))
        {
            return null;
        }

        var script = ResolveByNameOrIndex(input.Trim(), scripts);
        if (script is null)
        {
            Console.WriteLine($"Script '{input}' was not found.");
        }

        return script;
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

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  dotnet run --project AutomationTest -- --list");
        Console.WriteLine("  dotnet run --project AutomationTest -- --script <name|number>");
        Console.WriteLine("  dotnet run --project AutomationTest");
    }
}