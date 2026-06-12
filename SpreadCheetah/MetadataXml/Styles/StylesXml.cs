using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml.Styles;

internal static class StylesXml
{
    public static ValueTask WriteStylesAsync(
        this ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        StyleManager styleManager,
        CancellationToken token)
    {
        const string entryName = "xl/styles.xml";
        var namedStyles = styleManager.GetEmbeddedNamedStyles();
        var numberFormatCounter = new NumberFormatCounter();
        var writer = new StylesXmlWriter(styleManager, namedStyles, numberFormatCounter, buffer);
        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }
}

file struct StylesXmlWriter(
    StyleManager styleManager,
    List<(string, AddedStyle, StyleNameVisibility)>? namedStyles,
    NumberFormatCounter numberFormatCounter,
    SpreadsheetBuffer buffer)
    : IXmlWriter<StylesXmlWriter>
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""u8;

    private static ReadOnlySpan<byte> Footer => """</styleSheet>"""u8;

    private NumberFormatsXmlPart _numberFormatsXml = new(styleManager.UniqueCustomFormats?.GetList(), numberFormatCounter, buffer);
    private BordersXmlPart _bordersXml = new(styleManager.UniqueBorders?.GetList(), buffer);
    private FillsXmlPart _fillsXml = new(styleManager.UniqueFills?.GetList(), buffer);
    private FontXmlPart? _fontXml;
    private CellStylesXmlPart _cellStylesXml = new(namedStyles, buffer);
    private StyleXfXml? _currentXfXmlWriter;
    private StyleDxfXml? _currentDxfXmlWriter;
    private Element _next;
    private int _nextIndex;

    public readonly StylesXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => buffer.TryWrite(Header),
            Element.NumberFormats => _numberFormatsXml.TryWrite(),
            Element.FontsStart => TryWriteFontsStart(),
            Element.Fonts => TryWriteFonts(),
            Element.FontsEnd => buffer.TryWrite("</fonts>"u8),
            Element.Fills => _fillsXml.TryWrite(),
            Element.Borders => _bordersXml.TryWrite(),
            Element.CellStyleXfsStart => TryWriteCellStyleXfsStart(),
            Element.CellStyleXfsEntries => TryWriteCellStyleXfsEntries(),
            Element.CellXfsStart => TryWriteCellXfsStart(),
            Element.CellXfsEntries => TryWriteCellXfsEntries(),
            Element.CellXfsEnd => buffer.TryWrite("</cellXfs>"u8),
            Element.CellStyles => _cellStylesXml.TryWrite(),
            Element.DxfsStart => TryWriteDxfsStart(),
            Element.DxfsEntries => TryWriteDxfsEntries(),
            Element.DxfsEnd => TryWriteDxfsEnd(),
            _ => buffer.TryWrite(Footer),
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteFontsStart()
    {
        var fontCount = styleManager.UniqueFonts?.GetList().Count ?? 0;
        var defaultFont = styleManager.DefaultFont;
        var defaultFontSize = defaultFont?.Size ?? DefaultFont.DefaultSize;
        var defaultFontName = defaultFont?.Name ?? DefaultFont.DefaultName;
        Debug.Assert(defaultFontName.Length <= 31);

        return buffer.TryWrite(
            $"{"<fonts count=\""u8}" +
            $"{fontCount + 1}" +
            $"{"\"><font><sz val=\""u8}" +
            $"{defaultFontSize}" +
            $"{"\"/><name val=\""u8}" +
            $"{defaultFontName}" +
            $"{"\"/></font>"u8}");
    }

    private bool TryWriteFonts()
    {
        var fontsList = styleManager.UniqueFonts?.GetList();
        if (fontsList is null)
            return true;

        for (; _nextIndex < fontsList.Count; ++_nextIndex)
        {
            var font = fontsList[_nextIndex];
            Debug.Assert(font != ImmutableFont.From(styleManager.DefaultFont));

            var xml = _fontXml ?? new FontXmlPart(buffer, font);
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
        var count = (namedStyles?.Count ?? 0) + 1;
        return buffer.TryWrite(
            $"{"<cellStyleXfs count=\""u8}" +
            $"{count}" +
            $"{"\"><xf numFmtId=\"0\" fontId=\"0\"/>"u8}");
    }

    private bool TryWriteCellStyleXfsEntries()
    {
        if (namedStyles is not { } namedStylesLocal)
            return true;

        var alignments = styleManager.UniqueAlignments?.GetList();

        for (; _nextIndex < namedStylesLocal.Count; ++_nextIndex)
        {
            if (_currentXfXmlWriter is not { } writer)
            {
                var (_, addedStyle, _) = namedStylesLocal[_nextIndex];
                writer = new StyleXfXml(addedStyle, null, alignments, false, buffer);
            }

            if (!writer.TryWrite())
            {
                _currentXfXmlWriter = writer;
                return false;
            }

            _currentXfXmlWriter = null;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteCellXfsStart()
    {
        return buffer.TryWrite(
            $"{"</cellStyleXfs><cellXfs count=\""u8}" +
            $"{styleManager.AddedStyles.Count}" +
            $"{"\">"u8}");
    }

    private bool TryWriteCellXfsEntries()
    {
        var alignments = styleManager.UniqueAlignments?.GetList();
        var addedStyles = styleManager.AddedStyles;

        for (; _nextIndex < addedStyles.Count; ++_nextIndex)
        {
            if (_currentXfXmlWriter is not { } writer)
            {
                var addedStyle = addedStyles[_nextIndex];
                var embeddedNamedStyleIndex = GetEmbeddedNameStyleIndex(addedStyle);
                writer = new StyleXfXml(addedStyle, embeddedNamedStyleIndex, alignments, true, buffer);
            }

            if (!writer.TryWrite())
            {
                _currentXfXmlWriter = writer;
                return false;
            }

            _currentXfXmlWriter = null;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly int? GetEmbeddedNameStyleIndex(AddedStyle addedStyle)
    {
        if (addedStyle is not { Name: { } name, Visibility: not null })
            return null;
        if (styleManager.NamedStyles is not { } namedStylesLocal)
            return null;
        if (!namedStylesLocal.TryGetValue(name, out var value))
            return null;

        return value.NamedStyleIndex;
    }

    private readonly bool TryWriteDxfsStart()
    {
        var count = styleManager.DifferentialStyles?.Count ?? 0;
        return count == 0
            ? buffer.TryWrite("""<dxfs count="0"/>"""u8)
            : buffer.TryWrite($"{"<dxfs count=\""u8}{count}{"\">"u8}");
    }

    private bool TryWriteDxfsEntries()
    {
        if (styleManager.DifferentialStyles is not { } dxfsList)
            return true;

        for (; _nextIndex < dxfsList.Count; ++_nextIndex)
        {
            if (_currentDxfXmlWriter is not { } writer)
            {
                var dxfsEntry = dxfsList[_nextIndex];
                writer = new StyleDxfXml(dxfsEntry, numberFormatCounter, buffer);
            }

            if (!writer.TryWrite())
            {
                _currentDxfXmlWriter = writer;
                return false;
            }

            _currentDxfXmlWriter = null;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteDxfsEnd()
    {
        return styleManager.DifferentialStyles is not { } dxfsList
            || dxfsList.Count == 0
            || buffer.TryWrite("</dxfs>"u8);
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
        DxfsStart,
        DxfsEntries,
        DxfsEnd,
        Footer,
        Done
    }
}
