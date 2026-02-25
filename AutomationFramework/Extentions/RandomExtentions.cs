using System.Runtime.InteropServices.Swift;

namespace AutomationFramework.Extentions;

public static class Extentions
{
    public static double NextDouble(this Random random, double min, double max)
        => min + (random.NextDouble() * (max - min));

    public static float NextFloat(this Random random)
        => (float)random.NextDouble();

    public static float NextFloat(this Random random, float min, float max)
        => min + random.NextFloat() * (max - min);
}