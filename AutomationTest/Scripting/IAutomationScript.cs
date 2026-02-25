namespace AutomationTest.Scripting;

public interface IAutomationScript
{
    string Name { get; }

    string Description { get; }

    Task ExecuteAsync(ScriptExecutionContext context, CancellationToken cancellationToken);
}
