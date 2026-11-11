using SpreadCheetah.MetadataXml.Attributes;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.MetadataXml.Worksheets;

internal struct WorksheetColumnXmlPart(
    SpreadsheetBuffer buffer,
    int columnIndex,
    ColumnOptions options,
    StyleId? styleId,
    double? defaultColumnWidth)
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
            Element.Style => TryWriteStyle(),
            Element.Hidden => TryWriteHidden(),
            Element.OutlineLevel => TryWriteOutlineLevel(),
            Element.Collapsed => TryWriteCollapsed(),
            _ => buffer.TryWrite("/>"u8)
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
        if (options.Width is { } width)
        {
            return buffer.TryWrite(
                $"{" width=\""u8}" +
                $"{width}" +
                $"{"\" customWidth=\"1\""u8}");
        }

        if (options is { DefaultStyle: null, OutlineLevel: null, Collapsed: null })
            return true;

        return defaultColumnWidth is { } defaultWidth
            ? buffer.TryWrite($"{" width=\""u8}{defaultWidth}{"\""u8}")
            : buffer.TryWrite(" width=\"8.88671875\""u8);
    }

    private readonly bool TryWriteStyle()
    {
        return styleId is null
            || buffer.TryWrite($"{" style=\""u8}{styleId.Id}{"\""u8}");
    }

    private readonly bool TryWriteHidden()
    {
        return !options.Hidden || buffer.TryWrite(" hidden=\"1\""u8);
    }

    private readonly bool TryWriteOutlineLevel()
    {
        return options.OutlineLevel is null
            || buffer.TryWrite($"{new IntAttribute("outlineLevel"u8, options.OutlineLevel)}");
    }

    private readonly bool TryWriteCollapsed()
    {
        return options.Collapsed is null
            || buffer.TryWrite($"{new BooleanAttribute("collapsed"u8, options.Collapsed)}");
    }

    private enum Element
    {
        Header,
        Width,
        Style,
        Hidden,
        OutlineLevel,
        Collapsed,
        Footer,
        Done
    }
}
