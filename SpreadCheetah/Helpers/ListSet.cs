using SpreadCheetah.Styling;

namespace SpreadCheetah.Helpers;

internal sealed class ListSet<T> where T : notnull
{
    private readonly List<T> _list = [];
    private readonly Dictionary<T, int> _dictionary = [];

    public List<T> GetList() => _list;

    public int Add(in T item)
    {
        if (_dictionary.TryGetValue(item, out var index))
            return index;

        var newIndex = _list.Count;
        _list.Add(item);
        _dictionary[item] = newIndex;
        return newIndex;
    }
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