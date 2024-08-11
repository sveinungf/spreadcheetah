using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct FontsXmlPart(List<ImmutableFont> fonts, SpreadsheetBuffer buffer)
{
    private Element _next;
    private int _nextIndex;

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
            Element.Fonts => TryWriteFonts(),
            _ => buffer.TryWrite("</fonts>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        var span = buffer.GetSpan();
        var written = 0;

        const int defaultCount = 1;
        var totalCount = fonts.Count + defaultCount - 1;

        if (!"<fonts count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(totalCount, span, ref written)) return false;

        // The default font must be the first one (index 0)
        ReadOnlySpan<byte> defaultFont = "\">"u8 +
            """<font><sz val="11"/><name val="Calibri"/></font>"""u8;
        if (!defaultFont.TryCopyTo(span, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    private bool TryWriteFonts()
    {
        var defaultFont = new ImmutableFont { Size = Font.DefaultSize };
        var fontsLocal = fonts;

        for (; _nextIndex < fontsLocal.Count; ++_nextIndex)
        {
            var font = fonts[_nextIndex];
            if (font.Equals(defaultFont)) continue;

            var span = buffer.GetSpan();
            var written = 0;

            if (!"<font>"u8.TryCopyTo(span, ref written)) return false;
            if (font.Bold && !"<b/>"u8.TryCopyTo(span, ref written)) return false;
            if (font.Italic && !"<i/>"u8.TryCopyTo(span, ref written)) return false;
            if (font.Strikethrough && !"<strike/>"u8.TryCopyTo(span, ref written)) return false;
            if (!"<sz val=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(font.Size, span, ref written)) return false;
            if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;

            if (font.Color is { } color)
            {
                if (!"<color rgb=\""u8.TryCopyTo(span, ref written)) return false;
                if (!SpanHelper.TryWrite(color, span, ref written)) return false;
                if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;
            }

            var fontName = font.Name ?? "Calibri";
            if (!"<name val=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(fontName, span, ref written)) return false;
            if (!"\"/></font>"u8.TryCopyTo(span, ref written)) return false;

            buffer.Advance(written);
        }

        return true;
    }

    private enum Element
    {
        Header,
        Fonts,
        Footer,
        Done
    }
}
