using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Collections.Immutable;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct StylesXml
{
    private const string XmlPart2 =
        "</cellXfs>" +
        "<cellStyles count=\"1\">" +
        "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>" +
        "</cellStyles>" +
        "<dxfs count=\"0\"/>" +
        "</styleSheet>";

    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        CancellationToken token)
    {
        var stream = archive.CreateEntry("xl/styles.xml", compressionLevel).Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            var writer = new StylesXml(styles.Keys.ToList());
            var done = false;

            do
            {
                done = writer.TryWrite(buffer.GetSpan(), out var bytesWritten);
                buffer.Advance(bytesWritten);
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            } while (!done);

            await WriteAsync(stream, buffer, styles, writer.CustomNumberFormats, writer.Borders, writer.Fills, writer.Fonts, token).ConfigureAwait(false);
        }
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""u8;

    private readonly List<ImmutableStyle> _styles;
    private readonly StyleNumberFormatsXml _numberFormatsXml;
    private readonly StyleBordersXml _bordersXml;
    private readonly StyleFillsXml _fillsXml;
    private readonly StyleFontsXml _fontsXml;
    private Element _next;
    private int _nextIndex;

    // TODO: Make these fields
    public Dictionary<string, int>? CustomNumberFormats { get; }
    public Dictionary<ImmutableBorder, int> Borders { get; }
    public Dictionary<ImmutableFill, int> Fills { get; }
    public Dictionary<ImmutableFont, int> Fonts { get; }

    private StylesXml(List<ImmutableStyle> styles)
    {
        CustomNumberFormats = CreateCustomNumberFormatDictionary(styles);
        Borders = CreateBorderDictionary(styles);
        Fills = CreateFillDictionary(styles);
        Fonts = CreateFontDictionary(styles);
        _numberFormatsXml = new StyleNumberFormatsXml(CustomNumberFormats?.ToList());
        _bordersXml = new StyleBordersXml(Borders.Keys.ToList());
        _fillsXml = new StyleFillsXml(Fills.Keys.ToList());
        _fontsXml = new StyleFontsXml(Fonts.Keys.ToList());
        _styles = styles;
    }

    private static Dictionary<string, int>? CreateCustomNumberFormatDictionary(List<ImmutableStyle> styles)
    {
        Dictionary<string, int>? dictionary = null;
        var numberFormatId = 165; // Custom formats start sequentially from this ID

        foreach (var style in styles)
        {
            var numberFormat = style.NumberFormat;
            if (numberFormat is null) continue;
            if (NumberFormats.GetPredefinedNumberFormatId(numberFormat) is not null) continue;

            dictionary ??= new Dictionary<string, int>(StringComparer.Ordinal);
            if (dictionary.TryAdd(numberFormat, numberFormatId))
                ++numberFormatId;
        }

        return dictionary;
    }

    private static Dictionary<ImmutableBorder, int> CreateBorderDictionary(List<ImmutableStyle> styles)
    {
        var defaultBorder = new ImmutableBorder();
        const int defaultCount = 1;

        var uniqueBorders = new Dictionary<ImmutableBorder, int> { { defaultBorder, 0 } };
        var borderIndex = defaultCount;

        foreach (var style in styles)
        {
            var border = style.Border;
            if (!uniqueBorders.ContainsKey(border)) // TODO: Use CollectionsMarshal
            {
                uniqueBorders[border] = borderIndex;
                ++borderIndex;
            }
        }

        return uniqueBorders;
    }

    private static Dictionary<ImmutableFill, int> CreateFillDictionary(List<ImmutableStyle> styles)
    {
        var defaultFill = new ImmutableFill();
        const int defaultCount = 2;

        var uniqueFills = new Dictionary<ImmutableFill, int> { { defaultFill, 0 } };
        var fillIndex = defaultCount;

        foreach (var style in styles)
        {
            var fill = style.Fill;
            if (!uniqueFills.ContainsKey(fill)) // TODO: Use CollectionsMarshal
            {
                uniqueFills[fill] = fillIndex;
                ++fillIndex;
            }
        }

        return uniqueFills;
    }

    private static Dictionary<ImmutableFont, int> CreateFontDictionary(List<ImmutableStyle> styles)
    {
        var defaultFont = new ImmutableFont { Size = Font.DefaultSize };
        const int defaultCount = 1;

        var uniqueFonts = new Dictionary<ImmutableFont, int> { { defaultFont, 0 } };
        var fontIndex = defaultCount;

        foreach (var style in styles)
        {
            var font = style.Font;
            if (!uniqueFonts.ContainsKey(font)) // TODO: Use CollectionsMarshal
            {
                uniqueFonts[font] = fontIndex;
                ++fontIndex;
            }
        }

        return uniqueFonts;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.NumberFormats && !Advance(_numberFormatsXml.TryWrite(bytes, ref bytesWritten))) return false;
        if (_next == Element.Fonts && !Advance(_fontsXml.TryWrite(bytes, ref bytesWritten))) return false;
        if (_next == Element.Fills && !Advance(_fillsXml.TryWrite(bytes, ref bytesWritten))) return false;
        if (_next == Element.Borders && !Advance(_bordersXml.TryWrite(bytes, ref bytesWritten))) return false;
        if (_next == Element.CellXfsStart && !Advance(TryWriteCellXfsStart(bytes, ref bytesWritten))) return false;
        if (_next == Element.Styles && !Advance(TryWriteStyles(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteCellXfsStart(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        ReadOnlySpan<byte> xml =
            """<cellStyleXfs count="1"><xf numFmtId="0" fontId="0"/></cellStyleXfs>"""u8 +
            "<cellXfs count=\""u8;

        if (!xml.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_styles.Count, span, ref written)) return false;
        if (!"\">"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private bool TryWriteStyles(Span<byte> bytes, ref int bytesWritten)
    {
        var styles = _styles;

        for (; _nextIndex < styles.Count; ++_nextIndex)
        {
            var style = _styles[_nextIndex];
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            var numberFormatId = GetNumberFormatId(style.NumberFormat, CustomNumberFormats);

            if (!"<xf numFmtId=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(numberFormatId, span, ref written)) return false;
            if (!"\""u8.TryCopyTo(span, ref written)) return false;
            if (numberFormatId > 0 && !" applyNumberFormat=\"1\""u8.TryCopyTo(span, ref written)) return false;

            var fontIndex = Fonts.GetValueOrDefault(style.Font);
            if (!" fontId=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(fontIndex, span, ref written)) return false;
            if (!"\""u8.TryCopyTo(span, ref written)) return false;
            if (fontIndex > 0 && !" applyFont=\"1\""u8.TryCopyTo(span, ref written)) return false;

            var fillIndex = Fills.GetValueOrDefault(style.Fill);
            if (!" fillId=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(fillIndex, span, ref written)) return false;
            if (!"\""u8.TryCopyTo(span, ref written)) return false;
            if (fillIndex > 0 && !" applyFill=\"1\""u8.TryCopyTo(span, ref written)) return false;

            if (Borders.TryGetValue(style.Border, out var borderIndex) && borderIndex > 0)
            {
                if (!" borderId=\""u8.TryCopyTo(span, ref written)) return false;
                if (!SpanHelper.TryWrite(borderIndex, span, ref written)) return false;
                if (!"\" applyBorder=\"1\""u8.TryCopyTo(span, ref written)) return false;
            }

            if (!" xfId=\"0\""u8.TryCopyTo(span, ref written)) return false;

            var defaultAlignment = new ImmutableAlignment();
            if (style.Alignment == defaultAlignment)
            {
                if (!"/>"u8.TryCopyTo(span, ref written)) return false;
            }
            else
            {
                if (!""" applyAlignment="1">"""u8.TryCopyTo(span, ref written)) return false;
                if (!TryWriteAlignment(style.Alignment, span, ref written)) return false;
                if (!"</xf>"u8.TryCopyTo(span, ref written)) return false;
            }

            bytesWritten += written;
        }

        return true;
    }

    private static async ValueTask WriteAsync(
        Stream stream,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        Dictionary<string, int>? customNumberFormatLookup,
        Dictionary<ImmutableBorder, int> borderLookup,
        Dictionary<ImmutableFill, int> fillLookup,
        Dictionary<ImmutableFont, int> fontLookup,
        CancellationToken token)
    {
        await buffer.WriteAsciiStringAsync(XmlPart2, stream, token).ConfigureAwait(false);
        await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
    }

    private static int GetNumberFormatId(string? numberFormat, IReadOnlyDictionary<string, int>? customNumberFormats)
    {
        if (numberFormat is null) return 0;
        return NumberFormats.GetPredefinedNumberFormatId(numberFormat)
            ?? customNumberFormats?.GetValueOrDefault(numberFormat)
            ?? 0;
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

    private enum Element
    {
        Header,
        NumberFormats,
        Fonts,
        Fills,
        Borders,
        CellXfsStart,
        Styles,
        Done
    }
}
