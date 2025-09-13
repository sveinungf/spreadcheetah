using SpreadCheetah.Styling;

namespace SpreadCheetah.Helpers;

internal sealed class ListSet<T>
{
    private readonly List<T> _list = [];

    public List<T> GetList() => _list;

    public int Add(in T item)
    {
        var index = _list.IndexOf(item);
        if (index != -1)
            return index;

        _list.Add(item);
        return _list.Count - 1;
    }
}

internal readonly record struct AddedStyle(
    int? AlignmentIndex,
    int? BorderIndex,
    int? FillIndex,
    int? FontIndex,
    int? CustomFormatIndex,
    StandardNumberFormat? StandardFormat);