using System.Collections.Immutable;

namespace SpreadCheetah.Helpers;

internal static class ImmutableListExtensions
{
    public static int IndexOfOrDefault<T>(this ImmutableList<T> list, in T value, int defaultValue)
    {
        var index = list.IndexOf(value);
        return index == -1 ? defaultValue : index;
    }
}
