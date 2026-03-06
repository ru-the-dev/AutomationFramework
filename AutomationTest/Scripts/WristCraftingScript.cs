using System.Numerics;
using AutomationFramework.Extentions;
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
            await cursor.MoveToAsync(tailoringButtonPos, TimeSpan.FromMilliseconds(500), cancellationToken);
            cursor.Click();
            await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
            await cursor.MoveToAsync(createAllPos, TimeSpan.FromMilliseconds(500), cancellationToken);
            cursor.Click();
            await Task.Delay(TimeSpan.FromSeconds(Random.Shared.Next(130, 170)), cancellationToken);
            
            
            await cursor.MoveToAsync(closeTailoringButtonPos, TimeSpan.FromMilliseconds(500), cancellationToken);
            cursor.Click();
            await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
            await cursor.MoveToAsync(mountButtonPos, TimeSpan.FromMilliseconds(500), cancellationToken);
            cursor.Click();
            await Task.Delay(TimeSpan.FromSeconds(4), cancellationToken);
            await cursor.MoveToAsync(mailNPCPos, TimeSpan.FromMilliseconds(500), cancellationToken);
            cursor.Click(AutomationFramework.MouseButton.Right);
            await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
            await cursor.MoveToAsync(groupSectionPos, TimeSpan.FromMilliseconds(500), cancellationToken);
            cursor.Click();
            await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
            await cursor.MoveToAsync(sendMailButtonPos, TimeSpan.FromMilliseconds(500), cancellationToken);
            cursor.Click();
            await Task.Delay(TimeSpan.FromSeconds(Random.Shared.Next(10, 15)), cancellationToken);
            await cursor.MoveToAsync(closeMailButtonPos, TimeSpan.FromMilliseconds(500), cancellationToken);
            cursor.Click();
            await Task.Delay(TimeSpan.FromSeconds(Random.Shared.Next(5, 10)), cancellationToken);           
        }


        
    }


    private Vector2 GetRandomPointInBounds(RectangleF bounds)
    {
        var x = (float)(bounds.X + Random.Shared.NextFloat(0f, bounds.Width));
        var y = (float)(bounds.Top + Random.Shared.NextFloat(0f, bounds.Height));
        return new Vector2(x, y);
    }
}
