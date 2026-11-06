using Polyfills;

namespace SpreadCheetah.Test.Helpers;

internal static class RandomExtensions
{
    public static DateTime NextDateTime(this Random r, DateTime minValue, DateTime maxValue)
    {
        var minTicks = minValue.Ticks;
        var maxTicks = maxValue.Ticks;
        var ticks = r.NextInt64(minTicks, maxTicks);
        return new DateTime(ticks, DateTimeKind.Unspecified);
    }
}
