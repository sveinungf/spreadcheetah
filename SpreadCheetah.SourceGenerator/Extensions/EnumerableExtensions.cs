namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class EnumerableExtensions
{
    public static IEnumerable<(int Index, T Element)> Index<T>(this IEnumerable<T> elements)
    {
        return elements.Select((x, i) => (i, x));
    }
}
