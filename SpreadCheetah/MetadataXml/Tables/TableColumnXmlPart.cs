using SpreadCheetah.Helpers;
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
    private BufferWriteProgress _progress;

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
            Element.Name => TryWriteName(),
            Element.Label => TryWriteLabel(),
            Element.Function => TryWriteFunction(),
            _ => buffer.TryWrite("/>"u8)
        };

        if (Current)
        {
            _progress = default;
            ++_next;
        }

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        return buffer.TryWrite($"{"<tableColumn id=\""u8}{columnIndex + 1}{"\""u8}");
    }

    private bool TryWriteName()
    {
        return buffer.TryWrite(
            _progress, out _progress,
            $"{" name=\""u8}{headerName}{"\""u8}");
    }

    private bool TryWriteLabel()
    {
        if (totalRowLabel is null)
            return true;

        return buffer.TryWrite(
            _progress, out _progress,
            $"{" totalsRowLabel=\""u8}{totalRowLabel}{"\""u8}");
    }

    private readonly bool TryWriteFunction()
    {
        if (totalRowFunction is not { } function)
            return true;

        var functionAttributeValue = GetFunctionAttributeValue(function);
        return buffer.TryWrite($"{" totalsRowFunction=\""u8}{functionAttributeValue}{"\""u8}");
    }

    private static ReadOnlySpan<byte> GetFunctionAttributeValue(TableTotalRowFunction function) => function switch
    {
        TableTotalRowFunction.Average => "average"u8,
        TableTotalRowFunction.Count => "count"u8,
        TableTotalRowFunction.CountNumbers => "countNums"u8,
        TableTotalRowFunction.Maximum => "max"u8,
        TableTotalRowFunction.Minimum => "min"u8,
        TableTotalRowFunction.StandardDeviation => "stdDev"u8,
        TableTotalRowFunction.Sum => "sum"u8,
        _ => "var"u8
    };

    private enum Element
    {
        Header,
        Name,
        Label,
        Function,
        Footer,
        Done
    }
}
