namespace System.Collections.Generic;

#if NETSTANDARD2_0
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
}
#endif