namespace SpreadCheetah.Helpers;

internal sealed class OrderedSet<T> where T : notnull
{
#if NET9_0_OR_GREATER
    private readonly OrderedDictionary<T, byte> _dictionary = [];

    public IList<T> GetList() => _dictionary.Keys;

    public int Add(in T item)
    {
#if NET10_0_OR_GREATER
        _dictionary.TryAdd(item, 0, out var index);
        return index;
#else
        return _dictionary.TryAdd(item, 0)
            ? _dictionary.Count - 1
            : _dictionary.IndexOf(item);
#endif
    }
#else
    private readonly List<T> _list = [];
    private readonly Dictionary<T, int> _dictionary = [];

    public IList<T> GetList() => _list;

    public int Add(in T item)
    {
        if (_dictionary.TryGetValue(item, out var index))
            return index;

        var newIndex = _list.Count;
        _list.Add(item);
        _dictionary[item] = newIndex;
        return newIndex;
    }
#endif
}
