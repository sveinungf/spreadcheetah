using System.Globalization;

namespace SpreadCheetah.Test.Helpers;

internal static class ObjectExtensions
{
    public static string? GetExpectedCachedValueAsString(this object? cachedValue) => cachedValue switch
    {
        bool boolean => boolean.ToString().ToUpperInvariant(),
        _ => Convert.ToString(cachedValue, CultureInfo.InvariantCulture)
    };
}
