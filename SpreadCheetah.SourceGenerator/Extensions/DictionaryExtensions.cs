namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class DictionaryExtensions
{
    public static void AddWithImplicitKeys<T>(
        this Dictionary<int, T> dictionary,
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
