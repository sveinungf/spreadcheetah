using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct FontXmlPart(
    SpreadsheetBuffer buffer,
    ImmutableFont font)
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
            Element.Color => TryWriteColor(),
            _ => TryWriteName()
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        var bold = font.Bold ? "<b/>"u8 : ""u8;
        var italic = font.Italic ? "<i/>"u8 : ""u8;
        var strikethrough = font.Strikethrough ? "<strike/>"u8 : ""u8;
        var underline = GetUnderlineElement(font.Underline);

        return buffer.TryWrite(
            $"{"<font>"u8}" +
            $"{bold}{italic}{strikethrough}{underline}" +
            $"{"<sz val=\""u8}{font.Size}{"\"/>"u8}");
    }

    private static ReadOnlySpan<byte> GetUnderlineElement(Underline underline) => underline switch
    {
        Underline.Single => "<u/>"u8,
        Underline.SingleAccounting => """<u val="singleAccounting"/>"""u8,
        Underline.Double => """<u val="double"/>"""u8,
        Underline.DoubleAccounting => """<u val="doubleAccounting"/>"""u8,
        _ => []
    };

    private readonly bool TryWriteColor()
    {
        if (font.Color is not { } color)
            return true;

        return buffer.TryWrite(
            $"{"<color rgb=\""u8}" +
            $"{color}" +
            $"{"\"/>"u8}");
    }

    private readonly bool TryWriteName()
    {
        return buffer.TryWrite(
            $"{"<name val=\""u8}" +
            $"{font.Name}" +
            $"{"\"/></font>"u8}");
    }

    private enum Element
    {
        Header,
        Color,
        Name,
        Done
    }
}
