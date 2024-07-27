using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.IO.Compression;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct StylesXml
{
    public static async ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        Dictionary<ImmutableStyle, int> styles,
        Dictionary<string, (StyleId, StyleNameVisibility?)>? namedStyles,
        CancellationToken token)
    {
        var entry = archive.CreateEntry("xl/styles.xml", compressionLevel);
        var stream = entry.Open();
#if NETSTANDARD2_0
        using (stream)
#else
        await using (stream.ConfigureAwait(false))
#endif
        {
            // The order of Dictionary.Keys is not guaranteed, so we make sure the styles are sorted by the StyleId here.
            var orderedStyles = styles.OrderBy(x => x.Value).Select(x => x.Key).ToList();

            var filteredNamedStyles = namedStyles?
                .Where(x => x.Value.Item2 is not null)
                .Select(x => (x.Key, x.Value.Item1, x.Value.Item2.GetValueOrDefault()))
                .ToList();

            var writer = new StylesXml(orderedStyles, filteredNamedStyles, buffer);

            foreach (var success in writer)
            {
                if (!success)
                    await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            }

            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""u8;

    private static ReadOnlySpan<byte> Footer =>
        """<dxfs count="0"/>"""u8 +
        """</styleSheet>"""u8;

    private readonly List<ImmutableStyle> _styles;
    private readonly Dictionary<string, int>? _customNumberFormats;
    private readonly Dictionary<ImmutableBorder, int> _borders;
    private readonly Dictionary<ImmutableFill, int> _fills;
    private readonly Dictionary<ImmutableFont, int> _fonts;
    private readonly SpreadsheetBuffer _buffer;
    private NumberFormatsXmlPart _numberFormatsXml;
    private BordersXmlPart _bordersXml;
    private FillsXmlPart _fillsXml;
    private FontsXmlPart _fontsXml;
    private CellStylesXmlPart _cellStylesXml;
    private Element _next;
    private int _nextIndex;

    private StylesXml(
        List<ImmutableStyle> styles,
        List<(string, StyleId, StyleNameVisibility)>? namedStyles,
        SpreadsheetBuffer buffer)
    {
        _customNumberFormats = CreateCustomNumberFormatDictionary(styles);
        _borders = CreateBorderDictionary(styles);
        _fills = CreateFillDictionary(styles);
        _fonts = CreateFontDictionary(styles);
        _buffer = buffer;
        _numberFormatsXml = new NumberFormatsXmlPart(_customNumberFormats is { } formats ? [.. formats] : null, buffer);
        _bordersXml = new BordersXmlPart([.. _borders.Keys], buffer);
        _fillsXml = new FillsXmlPart([.. _fills.Keys], buffer);
        _fontsXml = new FontsXmlPart([.. _fonts.Keys], buffer);
        _cellStylesXml = new CellStylesXmlPart(namedStyles, buffer);
        _styles = styles;
    }

    public readonly StylesXml GetEnumerator() => this;
    public bool Current { get; private set; }

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

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.NumberFormats => _numberFormatsXml.TryWrite(),
            Element.Fonts => _fontsXml.TryWrite(),
            Element.Fills => _fillsXml.TryWrite(),
            Element.Borders => _bordersXml.TryWrite(),
            Element.CellXfsStart => TryWriteCellXfsStart(),
            Element.Styles => TryWriteStyles(),
            Element.CellXfsEnd => _buffer.TryWrite("</cellXfs>"u8),
            Element.CellStyles => _cellStylesXml.TryWrite(),
            _ => _buffer.TryWrite(Footer),
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteCellXfsStart()
    {
        var span = _buffer.GetSpan();
        var written = 0;

        ReadOnlySpan<byte> xml =
            """<cellStyleXfs count="1"><xf numFmtId="0" fontId="0"/></cellStyleXfs>"""u8 +
            "<cellXfs count=\""u8;

        if (!xml.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_styles.Count, span, ref written)) return false;
        if (!"\">"u8.TryCopyTo(span, ref written)) return false;

        _buffer.Advance(written);
        return true;
    }

    private bool TryWriteStyles()
    {
        var xfXml = new XfXmlPart(_buffer, _customNumberFormats, _borders, _fills, _fonts, true);
        var styles = _styles;

        for (; _nextIndex < styles.Count; ++_nextIndex)
        {
            var style = _styles[_nextIndex];
            if (!xfXml.TryWrite(style))
                return false;
        }

        return true;
    }

    private enum Element
    {
        Header,
        NumberFormats,
        Fonts,
        Fills,
        Borders,
        CellXfsStart,
        Styles,
        CellXfsEnd,
        CellStyles,
        Footer,
        Done
    }
}
