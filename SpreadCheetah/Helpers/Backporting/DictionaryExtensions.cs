#if NETSTANDARD2_0
namespace System.Collections.Generic;

internal static class DictionaryExtensions
{
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