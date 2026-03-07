namespace AutomationRunner.Scripting;


public abstract class BaseScript : IAutomationScript
{
    public abstract string Name { get; }

    public abstract string Description { get; }

    private bool _initialized = false;
    
    public async Task ExecuteAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        TimeSpan startDelay = TimeSpan.FromSeconds(2);
        Console.WriteLine($"Starting script: {Name} in {startDelay.TotalSeconds} seconds...");
        await Task.Delay(startDelay, cancellationToken);

        if (_initialized == false)
        {
            await InitializeAsync(context, cancellationToken);
            _initialized = true;
        }

        await RunAsync(context, cancellationToken); 
    }


    protected abstract Task InitializeAsync(ScriptExecutionContext context, CancellationToken cancellationToken);

    protected abstract Task RunAsync(ScriptExecutionContext context, CancellationToken cancellationToken);
    
    public abstract void Dispose();
}