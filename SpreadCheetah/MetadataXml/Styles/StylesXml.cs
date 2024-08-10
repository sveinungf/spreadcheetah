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
        StyleManager styleManager,
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
            var orderedStyles = styleManager.GetOrderedStyles();
            var embeddedNamedStyles = styleManager.GetEmbeddedNamedStyles();
            var writer = new StylesXml(orderedStyles, embeddedNamedStyles, buffer);

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
    private readonly List<(string, ImmutableStyle, StyleNameVisibility)>? _namedStyles;
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
        List<(string, ImmutableStyle, StyleNameVisibility)>? namedStyles,
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
        _namedStyles = namedStyles;
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
            Element.CellStyleXfsStart => TryWriteCellStyleXfsStart(),
            Element.CellStyleXfsEntries => TryWriteCellStyleXfsEntries(),
            Element.CellXfsStart => TryWriteCellXfsStart(),
            Element.CellXfsEntries => TryWriteCellXfsEntries(),
            Element.CellXfsEnd => _buffer.TryWrite("</cellXfs>"u8),
            Element.CellStyles => _cellStylesXml.TryWrite(),
            _ => _buffer.TryWrite(Footer),
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteCellStyleXfsStart()
    {
        var count = (_namedStyles?.Count ?? 0) + 1;
        return _buffer.TryWrite(
            $"{"<cellStyleXfs count=\""u8}" +
            $"{count}" +
            $"{"\"><xf numFmtId=\"0\" fontId=\"0\"/>"u8}");
    }

    private bool TryWriteCellStyleXfsEntries()
    {
        if (_namedStyles is not { } namedStyles)
            return true;

        var xfXml = new XfXmlPart(_buffer, _customNumberFormats, _borders, _fills, _fonts, false);

        for (; _nextIndex < _namedStyles.Count; ++_nextIndex)
        {
            var (_, style, _) = namedStyles[_nextIndex];
            if (!xfXml.TryWrite(style))
                return false;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteCellXfsStart()
    {
        return _buffer.TryWrite(
            $"{"</cellStyleXfs><cellXfs count=\""u8}" +
            $"{_styles.Count}" +
            $"{"\">"u8}");
    }

    private bool TryWriteCellXfsEntries()
    {
        var xfXml = new XfXmlPart(_buffer, _customNumberFormats, _borders, _fills, _fonts, true);
        var styles = _styles;

        for (; _nextIndex < styles.Count; ++_nextIndex)
        {
            // TODO: For cellXfs entries which are named styles, xfId should be the index into cellStyleXfs
            var style = _styles[_nextIndex];
            // TODO: Should pass xfId to this method somehow
            if (!xfXml.TryWrite(style))
                return false;
        }

        _nextIndex = 0;
        return true;
    }

    private enum Element
    {
        Header,
        NumberFormats,
        Fonts,
        Fills,
        Borders,
        CellStyleXfsStart,
        CellStyleXfsEntries,
        CellXfsStart,
        CellXfsEntries,
        CellXfsEnd,
        CellStyles,
        Footer,
        Done
    }
}
