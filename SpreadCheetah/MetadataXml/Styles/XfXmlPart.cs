using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.MetadataXml.Styles;

internal readonly struct XfXmlPart(
    SpreadsheetBuffer buffer,
    Dictionary<string, int>? customNumberFormats,
    bool cellXfsEntry)
{
    public bool TryWrite(in ImmutableStyle style, AddedStyle addedStyle, int? embeddedNamedStyleIndex)
    {
        var span = buffer.GetSpan();
        var written = 0;

        var numberFormatId = GetNumberFormatId(style.Format, customNumberFormats);

        if (!"<xf numFmtId=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(numberFormatId, span, ref written)) return false;
        if (!"\""u8.TryCopyTo(span, ref written)) return false;
        if (numberFormatId > 0 && !" applyNumberFormat=\"1\""u8.TryCopyTo(span, ref written)) return false;

        if (!TryWriteFont(addedStyle, span, ref written)) return false;
        if (!TryWriteFill(addedStyle, span, ref written)) return false;

        if (addedStyle.BorderIndex is { } borderIndex)
        {
            const int defaultBorders = 1;
            if (!" borderId=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(borderIndex + defaultBorders, span, ref written)) return false;
            if (!"\" applyBorder=\"1\""u8.TryCopyTo(span, ref written)) return false;
        }

        if (cellXfsEntry && !TryWriteXfId(embeddedNamedStyleIndex, span, ref written)) return false;

        if (style.Alignment == default)
        {
            if (!"/>"u8.TryCopyTo(span, ref written)) return false;
        }
        else
        {
            if (!""" applyAlignment="1">"""u8.TryCopyTo(span, ref written)) return false;
            if (!TryWriteAlignment(style.Alignment, span, ref written)) return false;
            if (!"</xf>"u8.TryCopyTo(span, ref written)) return false;
        }

        buffer.Advance(written);
        return true;
    }

    private static bool TryWriteXfId(int? embeddedNamedStyleIndex, Span<byte> span, ref int written)
    {
        var xfId = (embeddedNamedStyleIndex + 1) ?? 0;
        if (!" xfId=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(xfId, span, ref written)) return false;
        return "\""u8.TryCopyTo(span, ref written);
    }

    private static bool TryWriteFont(AddedStyle addedStyle, Span<byte> bytes, ref int bytesWritten)
    {
        const int defaultFonts = 1;
        var fontIndex = (addedStyle.FontIndex + defaultFonts) ?? 0;
        if (!" fontId=\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(fontIndex, bytes, ref bytesWritten)) return false;
        if (!"\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        return fontIndex == 0 || " applyFont=\"1\""u8.TryCopyTo(bytes, ref bytesWritten);
    }

    private static bool TryWriteFill(AddedStyle addedStyle, Span<byte> bytes, ref int bytesWritten)
    {
        const int defaultFills = 2;
        var fillIndex = (addedStyle.FillIndex + defaultFills) ?? 0;
        if (!" fillId=\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(fillIndex, bytes, ref bytesWritten)) return false;
        if (!"\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        return fillIndex == 0 || " applyFill=\"1\""u8.TryCopyTo(bytes, ref bytesWritten);
    }

    private static int GetNumberFormatId(NumberFormat? numberFormat, Dictionary<string, int>? customNumberFormats)
    {
        if (numberFormat is not { } format) return 0;
        if (format.StandardFormat is { } standardFormat) return (int)standardFormat;
        if (format.CustomFormat is null) return 0;
        return customNumberFormats?.GetValueOrDefault(format.CustomFormat) ?? 0;
    }

    private static bool TryWriteAlignment(ImmutableAlignment alignment, Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"<alignment"u8.TryCopyTo(span, ref written)) return false;
        if (!TryWriteHorizontalAlignment(alignment.Horizontal, span, ref written)) return false;
        if (!TryWriteVerticalAlignment(alignment.Vertical, span, ref written)) return false;
        if (alignment.WrapText && !" wrapText=\"1\""u8.TryCopyTo(span, ref written)) return false;

        if (alignment.Indent > 0)
        {
            if (!" indent=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(alignment.Indent, span, ref written)) return false;
            if (!"\""u8.TryCopyTo(span, ref written)) return false;
        }

        if (!"/>"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private static bool TryWriteHorizontalAlignment(HorizontalAlignment alignment, Span<byte> bytes, ref int bytesWritten) => alignment switch
    {
        HorizontalAlignment.Left => " horizontal=\"left\""u8.TryCopyTo(bytes, ref bytesWritten),
        HorizontalAlignment.Center => " horizontal=\"center\""u8.TryCopyTo(bytes, ref bytesWritten),
        HorizontalAlignment.Right => " horizontal=\"right\""u8.TryCopyTo(bytes, ref bytesWritten),
        _ => true
    };

    private static bool TryWriteVerticalAlignment(VerticalAlignment alignment, Span<byte> bytes, ref int bytesWritten) => alignment switch
    {
        VerticalAlignment.Center => " vertical=\"center\""u8.TryCopyTo(bytes, ref bytesWritten),
        VerticalAlignment.Top => " vertical=\"top\""u8.TryCopyTo(bytes, ref bytesWritten),
        _ => true
    };
}
