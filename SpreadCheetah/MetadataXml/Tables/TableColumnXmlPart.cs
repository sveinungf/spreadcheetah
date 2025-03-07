using SpreadCheetah.Tables;

namespace SpreadCheetah.MetadataXml.Tables;

internal struct TableColumnXmlPart(
    SpreadsheetBuffer buffer,
    int columnIndex,
    string headerName,
    string? totalRowLabel,
    TableTotalRowFunction? totalRowFunction)
{
    private Element _next;
    private int _currentStringIndex;

    public bool TryWrite()
    {
        while (MoveNext())
        {
            if (!Current)
                return false;
        }

        return true;
    }

    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.Name => TryWriteLongString(headerName),
            Element.LabelAttribute => TryWriteLabelAttribute(),
            Element.Label => TryWriteLabel(),
            _ => TryWriteFooter()
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        return buffer.TryWrite(
            $"{"<tableColumn id=\""u8}" +
            $"{columnIndex + 1}" +
            $"{"\" name=\""u8}");
    }

    private readonly bool TryWriteLabelAttribute()
    {
        return totalRowLabel is null || buffer.TryWrite("\" totalsRowLabel=\""u8);
    }

    private bool TryWriteLabel()
    {
        return totalRowLabel is null || TryWriteLongString(totalRowLabel);
    }

    private bool TryWriteLongString(string value)
    {
        if (buffer.WriteLongString(value, ref _currentStringIndex))
        {
            _currentStringIndex = 0;
            return true;
        }

        return false;
    }

    private readonly bool TryWriteFooter()
    {
        var functionAttribute = totalRowFunction is null ? [] : "\" totalsRowFunction=\""u8;
        var functionAttributeValue = GetFunctionAttributeValue(totalRowFunction);

        return buffer.TryWrite(
            $"{functionAttribute}" +
            $"{functionAttributeValue}" +
            $"{"\"/>"u8}");
    }

    private static ReadOnlySpan<byte> GetFunctionAttributeValue(TableTotalRowFunction? function) => function switch
    {
        TableTotalRowFunction.Average => "average"u8,
        TableTotalRowFunction.Count => "count"u8,
        TableTotalRowFunction.CountNumbers => "countNums"u8,
        TableTotalRowFunction.Maximum => "max"u8,
        TableTotalRowFunction.Minimum => "min"u8,
        TableTotalRowFunction.StandardDeviation => "stdDev"u8,
        TableTotalRowFunction.Sum => "sum"u8,
        TableTotalRowFunction.Variance => "var"u8,
        _ => []
    };

    private enum Element
    {
        Header,
        Name,
        LabelAttribute,
        Label,
        Footer,
        Done
    }
}
