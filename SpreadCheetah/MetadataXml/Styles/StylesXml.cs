using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct StylesXml : IXmlWriter<StylesXml>
{
    public static ValueTask WriteAsync(
        ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        StyleManager styleManager,
        CancellationToken token)
    {
        const string entryName = "xl/styles.xml";
        var writer = new StylesXml(styleManager, buffer);

        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }

    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""u8;

    private static ReadOnlySpan<byte> Footer =>
        """<dxfs count="0"/>"""u8 +
        """</styleSheet>"""u8;

    private readonly List<ImmutableFont> _fontsList;
    private readonly List<(string, AddedStyle, StyleNameVisibility)>? _namedStyles;
    private readonly StyleManager _styleManager;
    private readonly SpreadsheetBuffer _buffer;
    private NumberFormatsXmlPart _numberFormatsXml;
    private BordersXmlPart _bordersXml;
    private FillsXmlPart _fillsXml;
    private FontXmlPart? _fontXml;
    private CellStylesXmlPart _cellStylesXml;
    private Element _next;
    private int _nextIndex;

    private StylesXml(
        StyleManager styleManager,
        SpreadsheetBuffer buffer)
    {
        _styleManager = styleManager;
        _buffer = buffer;
        _numberFormatsXml = new NumberFormatsXmlPart(styleManager.UniqueCustomFormats.GetList(), buffer);
        _bordersXml = new BordersXmlPart(styleManager.UniqueBorders.GetList(), buffer);
        _fillsXml = new FillsXmlPart(styleManager.UniqueFills.GetList(), buffer);
        _fontsList = styleManager.UniqueFonts.GetList();
        _namedStyles = styleManager.GetEmbeddedNamedStyles();
        _cellStylesXml = new CellStylesXmlPart(_namedStyles, buffer);
    }

    public readonly StylesXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.NumberFormats => _numberFormatsXml.TryWrite(),
            Element.FontsStart => TryWriteFontsStart(),
            Element.Fonts => TryWriteFonts(),
            Element.FontsEnd => _buffer.TryWrite("</fonts>"u8),
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

    private readonly bool TryWriteFontsStart()
    {
        var defaultFont = _styleManager.DefaultFont;
        var defaultFontSize = defaultFont?.Size ?? DefaultFont.DefaultSize;
        var defaultFontName = defaultFont?.Name ?? DefaultFont.DefaultName;
        Debug.Assert(defaultFontName.Length <= 31);

        return _buffer.TryWrite(
            $"{"<fonts count=\""u8}" +
            $"{_fontsList.Count}" +
            $"{"\"><font><sz val=\""u8}" +
            $"{defaultFontSize}" +
            $"{"\"/><name val=\""u8}" +
            $"{defaultFontName}" +
            $"{"\"/></font>"u8}");
    }

    private bool TryWriteFonts()
    {
        var fontsLocal = _fontsList;

        for (; _nextIndex < fontsLocal.Count; ++_nextIndex)
        {
            var font = fontsLocal[_nextIndex];
            Debug.Assert(font != ImmutableFont.From(_styleManager.DefaultFont));

            var xml = _fontXml ?? new FontXmlPart(_buffer, font);
            if (!xml.TryWrite())
            {
                _fontXml = xml;
                return false;
            }

            _fontXml = null;
        }

        _nextIndex = 0;
        return true;
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

        var alignments = _styleManager.UniqueAlignments.GetList();
        var xfXml = new XfXmlPart(_buffer, alignments, false);

        for (; _nextIndex < _namedStyles.Count; ++_nextIndex)
        {
            var (_, addedStyle, _) = namedStyles[_nextIndex];
            if (!xfXml.TryWrite(addedStyle, null))
                return false;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteCellXfsStart()
    {
        return _buffer.TryWrite(
            $"{"</cellStyleXfs><cellXfs count=\""u8}" +
            $"{_styleManager.AddedStyles.Count}" +
            $"{"\">"u8}");
    }

    private bool TryWriteCellXfsEntries()
    {
        var alignments = _styleManager.UniqueAlignments.GetList();
        var xfXml = new XfXmlPart(_buffer, alignments, true);
        var addedStyles = _styleManager.AddedStyles;
        var namedStylesDictionary = _styleManager.NamedStyles;

        for (; _nextIndex < addedStyles.Count; ++_nextIndex)
        {
            var addedStyle = addedStyles[_nextIndex];
            int? embeddedNamedStyleIndex = addedStyle is { Name: { } name, Visibility: not null }
                && namedStylesDictionary is { } namedStyles && namedStyles.TryGetValue(name, out var value)
                ? value.NamedStyleIndex
                : null;

            if (!xfXml.TryWrite(addedStyle, embeddedNamedStyleIndex))
                return false;
        }

        _nextIndex = 0;
        return true;
    }

    private enum Element
    {
        Header,
        NumberFormats,
        FontsStart,
        Fonts,
        FontsEnd,
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
