using SpreadCheetah.ConditionalFormatting;
using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.Styling.Internal;

internal sealed class StyleManager
{
    private readonly NumberFormat? _defaultDateTimeFormat;

#if NET9_0_OR_GREATER
    private readonly OrderedDictionary<AddedStyle, StyleId> _styleDictionary = [];
    public IList<AddedStyle> AddedStyles => _styleDictionary.Keys;
#else
    private readonly Dictionary<AddedStyle, StyleId> _styleDictionary = [];
    private readonly List<AddedStyle> _addedStyles = [];
    public IList<AddedStyle> AddedStyles => _addedStyles;
#endif

    private void AddStyle(in AddedStyle style, StyleId styleId)
    {
#if NET9_0_OR_GREATER
        _styleDictionary.TryAdd(style, styleId);
#else
        if (_styleDictionary.TryAdd(style, styleId))
            _addedStyles.Add(style);
#endif
    }

    public bool HasStyles => _styleDictionary.Count > 0;

    public DefaultFont? DefaultFont { get; }
    public DefaultStyling? DefaultStyling { get; }

    public OrderedSet<ImmutableAlignment>? UniqueAlignments { get; private set; }
    public OrderedSet<ImmutableBorder>? UniqueBorders { get; private set; }
    public OrderedSet<ImmutableFill>? UniqueFills { get; private set; }
    public OrderedSet<ImmutableFont>? UniqueFonts { get; private set; }
    public OrderedSet<string>? UniqueCustomFormats { get; private set; }

    public List<AddedDifferentialStyle>? DifferentialStyles { get; private set; }
    public Dictionary<string, (StyleId StyleId, int NamedStyleIndex)>? NamedStyles { get; private set; }

    public StyleManager(NumberFormat? defaultDateTimeFormat, DefaultFont? defaultFont)
    {
        _defaultDateTimeFormat = defaultDateTimeFormat;
        DefaultFont = defaultFont;

        // If we have any style, the built-in default style must be the first one (meaning the first <xf> element in styles.xml).
        var defaultStyle = new AddedStyle();
        var styleId = AddStyleIfNotExists(defaultStyle);

        if (styleId.Id != styleId.DateTimeId)
            DefaultStyling = new DefaultStyling(styleId.DateTimeId);
    }

    public StyleId AddStyleIfNotExists(Style style)
    {
        var addedStyle = MapToAddedStyle(style, name: null, visibility: null);
        return AddStyleIfNotExists(addedStyle);
    }

    private StyleId AddStyleIfNotExists(in AddedStyle addedStyle)
    {
        if (_styleDictionary.TryGetValue(addedStyle, out var styleId))
            return styleId;

        var newId = _styleDictionary.Count;

        if (addedStyle.CustomFormatIndex is not null
            || addedStyle.StandardFormat is not null
            || _defaultDateTimeFormat is not { } dateTimeFormat)
        {
            styleId = new StyleId(newId, newId);
            AddStyle(addedStyle, styleId);
            return styleId;
        }

        var dateTimeStyle = addedStyle with
        {
            CustomFormatIndex = GetOrAddCustomFormat(dateTimeFormat.CustomFormat),
            Name = null,
            StandardFormat = dateTimeFormat.StandardFormat,
            Visibility = null
        };

        var existingDateTimeStyleId = _styleDictionary.GetValueOrDefault(dateTimeStyle);
        var dateTimeId = existingDateTimeStyleId?.Id ?? newId + 1;

        var newStyleId = new StyleId(newId, dateTimeId);
        AddStyle(addedStyle, newStyleId);
        AddStyle(dateTimeStyle, new StyleId(dateTimeId, dateTimeId));
        return newStyleId;
    }

    public int AddStyle(ConditionalFormatStyle style)
    {
        var differentialStyle = new AddedDifferentialStyle
        {
            Border = ImmutableBorder.From(style.GetBorderOrDefault()),
            Fill = ImmutableFill.From(style.GetFillOrDefault()),
            Font = ImmutableFont.From(style.GetFontOrDefault()),
            Format = style.Format
        };

        var differentialStyles = DifferentialStyles ??= [];
        differentialStyles.Add(differentialStyle);
        return differentialStyles.Count - 1;
    }

    private AddedStyle MapToAddedStyle(Style style, string? name, StyleNameVisibility? visibility)
    {
        return new AddedStyle(
            AlignmentIndex: GetAlignmentIndex(style),
            BorderIndex: GetBorderIndex(style),
            FillIndex: GetFillIndex(style),
            FontIndex: GetFontIndex(style),
            CustomFormatIndex: GetOrAddCustomFormat(style.Format?.CustomFormat),
            StandardFormat: style.Format?.StandardFormat,
            Name: name,
            Visibility: visibility);
    }

    private int? GetAlignmentIndex(Style style)
    {
        if (style.GetAlignmentOrDefault() is not { } styleAlignment)
            return null;

        var alignment = ImmutableAlignment.From(styleAlignment);
        if (alignment.IsDefault)
            return null;

        var uniqueAlignments = UniqueAlignments ??= new();
        return uniqueAlignments.Add(alignment);
    }

    private int? GetBorderIndex(Style style)
    {
        if (style.GetBorderOrDefault() is not { } styleBorder)
            return null;

        var border = ImmutableBorder.From(styleBorder);
        if (border.IsDefault)
            return null;

        var uniqueBorders = UniqueBorders ??= new();
        return uniqueBorders.Add(border);
    }

    private int? GetFillIndex(Style style)
    {
        if (style.GetFillOrDefault() is not { } styleFill)
            return null;

        var fill = ImmutableFill.From(styleFill);
        if (fill.IsDefault)
            return null;

        var uniqueFills = UniqueFills ??= new();
        return uniqueFills.Add(fill);
    }

    private int? GetFontIndex(Style style)
    {
        if (style.GetFontOrDefault() is not { } styleFont)
            return null;

        var font = ImmutableFont.From(styleFont, DefaultFont);
        if (font == ImmutableFont.From(DefaultFont))
            return null;

        var uniqueFonts = UniqueFonts ??= new();
        return uniqueFonts.Add(font);
    }

    private int? GetOrAddCustomFormat(string? customFormat)
    {
        if (customFormat is null)
            return null;

        var uniqueFormats = UniqueCustomFormats ??= new();
        return uniqueFormats.Add(customFormat);
    }

    public bool TryAddNamedStyle(string name, Style style, StyleNameVisibility? visibility,
        [NotNullWhen(true)] out StyleId? styleId)
    {
        styleId = null;

        var namedStyles = NamedStyles ??= new(StringComparer.OrdinalIgnoreCase);
        if (namedStyles.ContainsKey(name))
            return false;

        var addedStyle = MapToAddedStyle(style, name, visibility);
        styleId = AddStyleIfNotExists(addedStyle);

        // When adding a named style, we don't want to check for an existing style for potential reuse.
        // The reason is that if the named style would change, we only want it to affect parts where the
        // style name has been used explicitly.
        // Style names only refer to the main style, not the DateTime style.
        // Since the DateTime style doesn't have a name, it can be reused like regular styles.
        namedStyles[name] = (styleId, namedStyles.Count);
        return true;
    }

    public StyleId? GetStyleIdOrDefault(string name)
    {
        return NamedStyles is { } namedStyles && namedStyles.TryGetValue(name, out var value)
            ? value.StyleId
            : null;
    }

    public List<(string, AddedStyle, StyleNameVisibility)>? GetEmbeddedNamedStyles()
    {
        if (NamedStyles is null)
            return null;

        List<(string, AddedStyle, StyleNameVisibility)>? result = null;

        foreach (var addedStyle in AddedStyles)
        {
            if (addedStyle is not { Name: { } name, Visibility: { } visibility })
                continue;

            result ??= [];
            result.Add((name, addedStyle, visibility));
        }

        return result;
    }
}
