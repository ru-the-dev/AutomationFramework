using AutomationFramework;

namespace AutomationTest.Scripting;

public sealed class ScriptExecutionContext
{
    public AutomationFramework.Cursor Cursor { get; } = new();

    public Keyboard Keyboard { get; } = new();
}
