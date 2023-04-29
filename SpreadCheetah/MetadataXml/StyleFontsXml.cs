using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Buffers;

namespace SpreadCheetah.MetadataXml;

internal struct StyleFontsXml
{
    private readonly List<ImmutableFont> _fonts;
    private Element _next;
    private int _nextIndex;

    public StyleFontsXml(List<ImmutableFont> fonts)
    {
        _fonts = fonts;
    }

    public bool TryWrite(Span<byte> bytes, ref int bytesWritten)
    {
        if (_next == Element.Header && !Advance(TryWriteHeader(bytes, ref bytesWritten))) return false;
        if (_next == Element.Fonts && !Advance(TryWriteFonts(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance("</fonts>"u8.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteHeader(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        const int defaultCount = 1;
        var totalCount = _fonts.Count + defaultCount - 1;

        if (!"<fonts count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(totalCount, span, ref written)) return false;

        // The default font must be the first one (index 0)
        if (!"\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font>"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private bool TryWriteFonts(Span<byte> bytes, ref int bytesWritten)
    {
        var defaultFont = new ImmutableFont { Size = Font.DefaultSize };
        var fonts = _fonts;

        for (; _nextIndex < fonts.Count; ++_nextIndex)
        {
            var font = fonts[_nextIndex];
            if (font.Equals(defaultFont)) continue;

            var span = bytes.Slice(bytesWritten);
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
                if (!TryWriteColorChannel(color.A, span, ref written)) return false;
                if (!TryWriteColorChannel(color.R, span, ref written)) return false;
                if (!TryWriteColorChannel(color.G, span, ref written)) return false;
                if (!TryWriteColorChannel(color.B, span, ref written)) return false;
                if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;
            }

            var fontName = font.Name ?? "Calibri";
            if (!"<name val=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(fontName, span, ref written)) return false; // TODO: Handle very long font names, or limit font name length
            if (!"\"/></font>"u8.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        return true;
    }

    private static bool TryWriteColorChannel(int value, Span<byte> span, ref int written) => SpanHelper.TryWrite(value, span, ref written, new StandardFormat('X', 2));

    private enum Element
    {
        Header,
        Fonts,
        Footer,
        Done
    }
}
