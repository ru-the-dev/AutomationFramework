using System.Numerics;
using AutomationFramework.Extensions;
using AutomationRunner.Scripting;

namespace AutomationRunner.Scripts;

public sealed class DisenchantingLaptop1 : IAutomationScript
{
    public string Name => "disenchanting-laptop-1";

    public string Description => "Disenchants and opens mailbox";

    private RectangleF DisenchantButtonBounds = new RectangleF(659, 597, 225, 14); // Example bounds for the "Disenchant" button

    private RectangleF MailboxBounds = new RectangleF(133, 482, 100, 8); // Example bounds for the mailbox


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
                await cursor.MoveToAsync(target, cancellationToken: cancellationToken);
                await cursor.ClickAsync(cancellationToken: cancellationToken);

                // Wait a random time between clicks to mimic human behavior
                await Task.Delay(Random.Shared.Next(2500, 3500), cancellationToken);
            }

            // After clicking the "Disenchant" button several times, click the mailbox
            var mailboxTarget = GetRandomPointInBounds(MailboxBounds);
            await cursor.MoveToAsync(mailboxTarget, cancellationToken: cancellationToken);
            await cursor.ClickAsync(cancellationToken: cancellationToken);
        }


        
    }


    private Vector2 GetRandomPointInBounds(RectangleF bounds)
    {
        var x = (float)(bounds.X + Random.Shared.NextFloat(0f, bounds.Width));
        var y = (float)(bounds.Top + Random.Shared.NextFloat(0f, bounds.Height));
        return new Vector2(x, y);
    }
}
