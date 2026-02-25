namespace AutomationFramework;

public static class Duration
{
    private static readonly Random _random = new();

    /// <summary>
    /// Applies a random factor to the base duration to simulate human-like variability.
    /// </summary>
    public static TimeSpan ApplyRandomFactor(TimeSpan baseDuration, double minFactor, double maxFactor)
    {
        if (baseDuration <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(baseDuration), "Base duration must be greater than zero.");
        }

        if (minFactor <= 0 || maxFactor <= 0)
        {
            throw new ArgumentOutOfRangeException("Duration factors must be greater than zero.");
        }

        if (minFactor > maxFactor)
        {
            throw new ArgumentException("minFactor cannot be greater than maxFactor.");
        }

        double randomFactor = minFactor + (_random.NextDouble() * (maxFactor - minFactor));
        return TimeSpan.FromTicks((long)(baseDuration.Ticks * randomFactor));
    }
}