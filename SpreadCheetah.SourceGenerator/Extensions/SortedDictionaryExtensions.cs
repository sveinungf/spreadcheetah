namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class SortedDictionaryExtensions
{
    public static void AddWithImplicitKeys<T>(
        this SortedDictionary<int, T> dictionary,
        IEnumerable<T> values)
    {
        var implicitKey = 1;
        foreach (var value in values)
        {
            while (dictionary.ContainsKey(implicitKey))
                ++implicitKey;

            dictionary.Add(implicitKey, value);
            ++implicitKey;
        }
    }
}
