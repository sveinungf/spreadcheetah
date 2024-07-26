using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml;

internal struct StylesXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        CancellationToken token)
    {
        var entry = archive.CreateEntry("xl/styles.xml", compressionLevel);
        var writer = new StylesXml([.. styles.Keys]);
#pragma warning disable EPS06 // Hidden struct copy operation
        return writer.WriteAsync(entry, buffer, token);
#pragma warning restore EPS06 // Hidden struct copy operation
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""u8;

    private static ReadOnlySpan<byte> Footer =>
        """</cellXfs>"""u8 +
        """<cellStyles count="1">"""u8 +
        """<cellStyle name="Normal" xfId="0" builtinId="0"/>"""u8 +
        """</cellStyles>"""u8 +
        """<dxfs count="0"/>"""u8 +
        """</styleSheet>"""u8;

    private readonly List<ImmutableStyle> _styles;
    private readonly Dictionary<string, int>? _customNumberFormats;
    private readonly Dictionary<ImmutableBorder, int> _borders;
    private readonly Dictionary<ImmutableFill, int> _fills;
    private readonly Dictionary<ImmutableFont, int> _fonts;
    private StyleNumberFormatsXml _numberFormatsXml;
    private StyleBordersXml _bordersXml;
    private StyleFillsXml _fillsXml;
    private StyleFontsXml _fontsXml;
    private Element _next;
    private int _nextIndex;

    private StylesXml(List<ImmutableStyle> styles)
    {
        _customNumberFormats = CreateCustomNumberFormatDictionary(styles);
        _borders = CreateBorderDictionary(styles);
        _fills = CreateFillDictionary(styles);
        _fonts = CreateFontDictionary(styles);
        _numberFormatsXml = new StyleNumberFormatsXml(_customNumberFormats is { } formats ? [.. formats] : null);
        _bordersXml = new StyleBordersXml([.. _borders.Keys]);
        _fillsXml = new StyleFillsXml([.. _fills.Keys]);
        _fontsXml = new StyleFontsXml([.. _fonts.Keys]);
        _styles = styles;
    }

    private static Dictionary<string, int>? CreateCustomNumberFormatDictionary(List<ImmutableStyle> styles)
    {
        Dictionary<string, int>? dictionary = null;
        var numberFormatId = 165; // Custom formats start sequentially from this ID

        foreach (var style in styles)
        {
            var numberFormat = style.Format;
            if (numberFormat is not { } format) continue;
            if (format.CustomFormat is null) continue;

            dictionary ??= new Dictionary<string, int>(StringComparer.Ordinal);
            if (dictionary.TryAdd(format.CustomFormat, numberFormatId))
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
            if (uniqueBorders.TryAdd(style.Border, borderIndex))
                ++borderIndex;
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
            if (uniqueFills.TryAdd(style.Fill, fillIndex))
                ++fillIndex;
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
            if (uniqueFonts.TryAdd(style.Font, fontIndex))
                ++fontIndex;
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
        if (_next == Element.Footer && !Advance(Footer.TryCopyTo(bytes, ref bytesWritten))) return false;

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

            var numberFormatId = GetNumberFormatId(style.Format, _customNumberFormats);

            if (!"<xf numFmtId=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(numberFormatId, span, ref written)) return false;
            if (!"\""u8.TryCopyTo(span, ref written)) return false;
            if (numberFormatId > 0 && !" applyNumberFormat=\"1\""u8.TryCopyTo(span, ref written)) return false;

            if (!TryWriteFont(style.Font, span, ref written)) return false;
            if (!TryWriteFill(style.Fill, span, ref written)) return false;

            if (_borders.TryGetValue(style.Border, out var borderIndex) && borderIndex > 0)
            {
                if (!" borderId=\""u8.TryCopyTo(span, ref written)) return false;
                if (!SpanHelper.TryWrite(borderIndex, span, ref written)) return false;
                if (!"\" applyBorder=\"1\""u8.TryCopyTo(span, ref written)) return false;
            }

            if (!" xfId=\"0\""u8.TryCopyTo(span, ref written)) return false;

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

            bytesWritten += written;
        }

        return true;
    }

    private readonly bool TryWriteFont(ImmutableFont font, Span<byte> bytes, ref int bytesWritten)
    {
        var fontIndex = _fonts.GetValueOrDefault(font);
        if (!" fontId=\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(fontIndex, bytes, ref bytesWritten)) return false;
        if (!"\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        return fontIndex == 0 || " applyFont=\"1\""u8.TryCopyTo(bytes, ref bytesWritten);
    }

    private readonly bool TryWriteFill(ImmutableFill fill, Span<byte> bytes, ref int bytesWritten)
    {
        var fillIndex = _fills.GetValueOrDefault(fill);
        if (!" fillId=\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(fillIndex, bytes, ref bytesWritten)) return false;
        if (!"\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        return fillIndex == 0 || " applyFill=\"1\""u8.TryCopyTo(bytes, ref bytesWritten);
    }

    private static int GetNumberFormatId(NumberFormat? numberFormat, IReadOnlyDictionary<string, int>? customNumberFormats)
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

    private enum Element
    {
        Header,
        NumberFormats,
        Fonts,
        Fills,
        Borders,
        CellXfsStart,
        Styles,
        Footer,
        Done
    }
}
