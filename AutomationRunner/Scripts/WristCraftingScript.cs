using System.Numerics;
using AutomationFramework.Extensions;
using AutomationRunner.Scripting;

namespace AutomationRunner.Scripts;

public sealed class WristCraftingScript : IAutomationScript
{
    public string Name => "wrist-crafting-script";

    public string Description => "Crafts items using the wrist crafting method";

    public async Task ExecuteAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        var tailoringButtonPos = new Vector2(2264, 1309);
        var createAllPos = new Vector2(897, 913);
        var closeTailoringButtonPos = new Vector2(1176, 155);
        var mountButtonPos = new Vector2(1230, 1368);
        var mailNPCPos = new Vector2(1797, 833);
        var groupSectionPos = new Vector2(1342, 428);
        var sendMailButtonPos = new Vector2(1434, 987);
        var closeMailButtonPos = new Vector2(1753, 424);


        var cursor = new AutomationFramework.Cursor(); 

        
        while (true)
        {
            await cursor.MoveToAsync(tailoringButtonPos, cancellationToken: cancellationToken);
            await cursor.ClickAsync(cancellationToken: cancellationToken);
            await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
            await cursor.MoveToAsync(createAllPos, cancellationToken: cancellationToken);
            await cursor.ClickAsync(cancellationToken: cancellationToken);
            await Task.Delay(TimeSpan.FromSeconds(Random.Shared.Next(130, 170)), cancellationToken);
            
            
            await cursor.MoveToAsync(closeTailoringButtonPos, cancellationToken: cancellationToken);
            await cursor.ClickAsync(cancellationToken: cancellationToken);
            await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
            await cursor.MoveToAsync(mountButtonPos, cancellationToken: cancellationToken);
            await cursor.ClickAsync(cancellationToken: cancellationToken);
            await Task.Delay(TimeSpan.FromSeconds(4), cancellationToken);
            await cursor.MoveToAsync(mailNPCPos, cancellationToken: cancellationToken);
            await cursor.ClickAsync(AutomationFramework.MouseButton.Right, cancellationToken: cancellationToken);
            await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
            await cursor.MoveToAsync(groupSectionPos, cancellationToken: cancellationToken);
            await cursor.ClickAsync(cancellationToken: cancellationToken);
            await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
            await cursor.MoveToAsync(sendMailButtonPos, cancellationToken: cancellationToken);
            await cursor.ClickAsync(cancellationToken: cancellationToken);
            await Task.Delay(TimeSpan.FromSeconds(Random.Shared.Next(10, 15)), cancellationToken);
            await cursor.MoveToAsync(closeMailButtonPos, cancellationToken: cancellationToken);
            await cursor.ClickAsync(cancellationToken: cancellationToken);
            await Task.Delay(TimeSpan.FromSeconds(Random.Shared.Next(5, 10)), cancellationToken);           
        }
    }
}
