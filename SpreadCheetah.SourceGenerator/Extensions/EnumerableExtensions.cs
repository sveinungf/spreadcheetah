using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class EnumerableExtensions
{
    public static IEnumerable<(int Index, T Element)> Index<T>(this IEnumerable<T> elements)
    {
        return elements.Select((x, i) => (i, x));
    }

    public static EquatableArray<T> ToEquatableArray<T>(this IEnumerable<T> elements)
        where T : IEquatable<T>
    {
        var array = elements is T[] arr ? arr : elements.ToArray();
        return new EquatableArray<T>(array);
    }
}
