using System.Reflection;

namespace AutomationRunner.Scripting;

public static class ScriptRegistry
{
    public static IReadOnlyList<IAutomationScript> DiscoverScripts()
    {
        var scripts = Assembly
            .GetExecutingAssembly()
            .GetTypes()
            .Where(type => !type.IsAbstract && !type.IsInterface)
            .Where(type => typeof(IAutomationScript).IsAssignableFrom(type))
            .Where(type => type.GetConstructor(Type.EmptyTypes) is not null)
            .Select(type => (IAutomationScript)Activator.CreateInstance(type)!)
            .OrderBy(script => script.Name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var duplicateNames = scripts
            .GroupBy(script => script.Name, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .ToArray();

        if (duplicateNames.Length > 0)
        {
            throw new InvalidOperationException($"Duplicate script names found: {string.Join(", ", duplicateNames)}");
        }

        return scripts;
    }
}
