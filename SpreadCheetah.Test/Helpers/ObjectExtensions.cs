using System.Globalization;

namespace SpreadCheetah.Test.Helpers;

internal static class ObjectExtensions
{
    public static string? GetExpectedCachedValueAsString(this object? cachedValue) => cachedValue switch
    {
        DateTime dateTime => dateTime.ToOADate().ToString(),
        _ => Convert.ToString(cachedValue, CultureInfo.InvariantCulture)
    };
}
