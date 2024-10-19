namespace SpreadCheetah.SourceGenerator.Models;

internal sealed class StyleLookup
{
    private Dictionary<CellFormat, int>? _cellFormatToStyleIdIndex;
    private Dictionary<CellStyle, int>? _cellStyleToStyleIdIndex;

    private int StyleCount =>
        _cellFormatToStyleIdIndex?.Count ?? 0
        + _cellStyleToStyleIdIndex?.Count ?? 0;

    public bool TryAdd(CellFormat cellFormat)
    {
        var dictionary = _cellFormatToStyleIdIndex ??= [];
        return TryAdd(cellFormat, dictionary);
    }

    public bool TryAdd(CellStyle cellStyle)
    {
        var dictionary = _cellStyleToStyleIdIndex ??= [];
        return TryAdd(cellStyle, dictionary);
    }

    private bool TryAdd<T>(T key, Dictionary<T, int> dictionary)
    {
        if (dictionary.ContainsKey(key))
            return false;

        dictionary[key] = StyleCount;
        return true;
    }

    public bool TryGetStyleIdIndex(CellFormat cellFormat, out int? styleIdIndex)
    {
        return TryGetStyleIdIndex(cellFormat, _cellFormatToStyleIdIndex, out styleIdIndex);
    }

    public bool TryGetStyleIdIndex(CellStyle cellStyle, out int? styleIdIndex)
    {
        return TryGetStyleIdIndex(cellStyle, _cellStyleToStyleIdIndex, out styleIdIndex);
    }

    private static bool TryGetStyleIdIndex<T>(T key, Dictionary<T, int>? dictionary, out int? styleIdIndex)
    {
        styleIdIndex = null;
        if (dictionary is null)
            return false;

        if (!dictionary.TryGetValue(key, out var index))
            return false;

        styleIdIndex = index;
        return true;
    }
}
