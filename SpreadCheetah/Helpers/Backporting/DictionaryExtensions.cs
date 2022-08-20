#if NETSTANDARD2_0
namespace System.Collections.Generic;

internal static class DictionaryExtensions
{
    public static TValue? GetValueOrDefault<TKey, TValue>(this IReadOnlyDictionary<TKey, TValue> dictionary, TKey key)
    {
        return dictionary.GetValueOrDefault(key, default!);
    }

    public static TValue GetValueOrDefault<TKey, TValue>(this IReadOnlyDictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue)
    {
        if (dictionary is null)
            throw new ArgumentNullException(nameof(dictionary));

        return dictionary.TryGetValue(key, out TValue? value) ? value : defaultValue;
    }

    public static bool TryAdd<TKey, TValue>(this Dictionary<TKey, TValue> dict, TKey key, TValue value)
        where TKey : notnull
    {
        if (dict.ContainsKey(key))
            return false;

        dict.Add(key, value);
        return true;
    }
}
#endif