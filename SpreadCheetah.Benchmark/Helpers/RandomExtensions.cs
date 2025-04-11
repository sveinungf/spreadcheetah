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

    public static DateTime NextDateTime(this Random r, bool withFraction)
    {
        var origin = new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        return withFraction
            ? origin.AddSeconds(r.Next(-300_000_000, 300_000_000))
            : origin.AddDays(r.Next(-10000, 10000));
    }
}
