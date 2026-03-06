using System.Numerics;
using AutomationFramework.Extentions;
using AutomationRunner.Scripting;

namespace AutomationRunner.Scripts;

public sealed class DisenchantingScript1 : IAutomationScript
{
    public string Name => "disenchanting-script-1";

    public string Description => "Disenchants and opens mailbox";

    private RectangleF DisenchantButtonBounds = new RectangleF(1830, 868, 279, 21); // Example bounds for the "Disenchant" button

    private RectangleF MailboxBounds = new RectangleF(1298, 932, 282, 20); // Example bounds for the mailbox


    public async Task ExecuteAsync(ScriptExecutionContext context, CancellationToken cancellationToken)
    {
        var cursor = new AutomationFramework.Cursor(); 

        
        while (true)
        {
            int iterations = Random.Shared.Next(50, 100);

            for (int i = 0; i < iterations; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var target = GetRandomPointInBounds(DisenchantButtonBounds);
                await cursor.MoveToAsync(target, TimeSpan.FromMilliseconds(200), cancellationToken);
                cursor.Click();

                // Wait a random time between clicks to mimic human behavior
                await Task.Delay(Random.Shared.Next(2500, 3500), cancellationToken);
            }

            // After clicking the "Disenchant" button several times, click the mailbox
            var mailboxTarget = GetRandomPointInBounds(MailboxBounds);
            await cursor.MoveToAsync(mailboxTarget, TimeSpan.FromMilliseconds(200), cancellationToken);
            cursor.Click();
        }


        
    }


    private Vector2 GetRandomPointInBounds(RectangleF bounds)
    {
        var x = (float)(bounds.X + Random.Shared.NextFloat(0f, bounds.Width));
        var y = (float)(bounds.Top + Random.Shared.NextFloat(0f, bounds.Height));
        return new Vector2(x, y);
    }
}
