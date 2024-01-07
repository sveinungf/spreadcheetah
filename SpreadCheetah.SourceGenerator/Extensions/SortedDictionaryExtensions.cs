namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class SortedDictionaryExtensions
{
    public static void AddWithImplicitKeys(
        this SortedDictionary<int, string> dictionary,
        IEnumerable<string> values)
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
