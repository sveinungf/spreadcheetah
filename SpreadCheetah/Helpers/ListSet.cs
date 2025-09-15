using SpreadCheetah.Styling;

namespace SpreadCheetah.Helpers;

internal sealed class ListSet<T> where T : notnull
{
#if NET9_0_OR_GREATER
    private readonly OrderedDictionary<T, byte> _dictionary = [];

    public IList<T> GetList() => _dictionary.Keys;

    public int Add(in T item)
    {
        return _dictionary.TryAdd(item, 0)
            ? _dictionary.Count - 1
            : _dictionary.IndexOf(item);
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

internal readonly record struct AddedStyle(
    int? AlignmentIndex,
    int? BorderIndex,
    int? FillIndex,
    int? FontIndex,
    int? CustomFormatIndex,
    StandardNumberFormat? StandardFormat,
    string? Name,
    StyleNameVisibility? Visibility);