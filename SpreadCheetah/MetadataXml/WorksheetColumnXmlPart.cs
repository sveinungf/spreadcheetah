using SpreadCheetah.Worksheets;

namespace SpreadCheetah.MetadataXml;

internal struct WorksheetColumnXmlPart(
    SpreadsheetBuffer buffer,
    int columnIndex,
    ColumnOptions options)
{
    private Element _next;

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
            Element.Width => TryWriteWidth(),
            Element.Hidden => TryWriteHidden(),
            _ => buffer.TryWrite(" />"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        return buffer.TryWrite(
            $"{"<col min=\""u8}" +
            $"{columnIndex}" +
            $"{"\" max=\""u8}" +
            $"{columnIndex}" +
            $"{"\""u8}");
    }

    private readonly bool TryWriteWidth()
    {
        if (options.Width is not { } width)
            return true;

        return buffer.TryWrite(
            $"{" width=\""u8}" +
            $"{width}" +
            $"{"\" customWidth=\"1\""u8}");
    }

    private readonly bool TryWriteHidden()
    {
        return !options.Hidden || buffer.TryWrite(" hidden=\"1\""u8);
    }

    private enum Element
    {
        Header,
        Width,
        Hidden,
        Footer,
        Done
    }
}
