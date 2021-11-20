namespace SpreadCheetah.Benchmark.Helpers;

internal static class RandomExtensions
{
    public static decimal NextDecimal(this Random r, decimal minValue, decimal maxValue)
    {
        return (decimal)r.NextDouble() * (maxValue - minValue) + minValue;
    }

    public static bool NextBoolean(this Random r)
    {
        return r.Next() > (int.MaxValue / 2);
    }
}
