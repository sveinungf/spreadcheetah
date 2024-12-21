using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class EnumerableExtensions
{
    public static EquatableArray<T> ToEquatableArray<T>(this IEnumerable<T> elements)
        where T : IEquatable<T>
    {
        var array = elements is T[] arr ? arr : [.. elements];
        return new EquatableArray<T>(array);
    }
}
