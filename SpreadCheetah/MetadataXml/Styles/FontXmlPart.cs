using SpreadCheetah.Helpers;
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

        return buffer.TryWrite(
            $"{"<font>"u8}" +
            $"{bold}{italic}{strikethrough}" +
            $"{"<sz val=\""u8}{font.Size}{"\"/>"u8}");
    }

    private readonly bool TryWriteColor()
    {
        if (font.Color is not { } color)
            return true;

        // TODO: Add support for Color in buffer.TryWrite
        var written = 0;
        var span = buffer.GetSpan();
        if (!"<color rgb=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(color, span, ref written)) return false;
        if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    private readonly bool TryWriteName()
    {
        var fontName = font.Name ?? "Calibri";
        return buffer.TryWrite(
            $"{"<name val=\""u8}" +
            $"{fontName}" +
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
