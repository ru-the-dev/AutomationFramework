using Microsoft.Extensions.Configuration;

namespace AutomationTest.Scripting;

public sealed class ScriptExecutionContext
{
    public ScriptExecutionContext(IConfiguration configuration)
    {
        Configuration = configuration;
    }

    public IConfiguration Configuration { get; }
}
