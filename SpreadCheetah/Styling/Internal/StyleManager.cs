using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.Styling.Internal;

internal sealed class StyleManager
{
    private readonly NumberFormat? _defaultDateTimeFormat;
    private readonly Dictionary<ImmutableStyle, StyleId> _styleDictionary = [];

    public DefaultFont? DefaultFont { get; }
    public DefaultStyling? DefaultStyling { get; }
    public List<StyleElement> StyleElements { get; } = [];

    // TODO: Null if nothing added
    public ListSet<ImmutableAlignment> UniqueAlignments { get; } = new();
    public ListSet<ImmutableBorder> UniqueBorders { get; } = new();
    public ListSet<ImmutableFill> UniqueFills { get; } = new();
    public ListSet<ImmutableFont> UniqueFonts { get; } = new();
    public ListSet<string> UniqueCustomFormats { get; } = new();

    public List<AddedStyle> AddedStyles { get; } = [];

    public Dictionary<string, (StyleId StyleId, int NamedStyleIndex)>? NamedStyles { get; private set; }

    public StyleManager(NumberFormat? defaultDateTimeFormat, DefaultFont? defaultFont)
    {
        _defaultDateTimeFormat = defaultDateTimeFormat;
        DefaultFont = defaultFont;

        // If we have any style, the built-in default style must be the first one (meaning the first <xf> element in styles.xml).
        var defaultStyle = new ImmutableStyle() with { Font = ImmutableFont.From(defaultFont) };
        var styleId = AddStyleIfNotExists(defaultStyle);

        if (styleId.Id != styleId.DateTimeId)
            DefaultStyling = new DefaultStyling(styleId.DateTimeId);
    }

    public StyleId AddStyleIfNotExists(in ImmutableStyle style)
    {
        if (_styleDictionary.TryGetValue(style, out var id))
            return id;

        var newId = AssignStyleId(style, name: null, visibility: null);
        var dateTimeId = AssignDateTimeStyleId(style) ?? newId;
        var styleId = new StyleId(newId, dateTimeId);
        _styleDictionary[style] = styleId;
        return styleId;
    }

    private int AssignStyleId(in ImmutableStyle style, string? name, StyleNameVisibility? visibility)
    {
        int? alignmentIndex = style.Alignment != default
            ? UniqueAlignments.Add(style.Alignment)
            : null;

        int? borderIndex = style.Border != default
            ? UniqueBorders.Add(style.Border)
            : null;

        int? fillIndex = style.Fill != default
            ? UniqueFills.Add(style.Fill)
            : null;

        int? fontIndex = style.Font != ImmutableFont.From(DefaultFont)
            ? UniqueFonts.Add(style.Font)
            : null;

        int? formatIndex = style.Format is { CustomFormat: { } format }
            ? UniqueCustomFormats.Add(format)
            : null;

        var addedStyle = new AddedStyle(
            AlignmentIndex: alignmentIndex,
            BorderIndex: borderIndex,
            FillIndex: fillIndex,
            FontIndex: fontIndex,
            CustomFormatIndex: formatIndex,
            StandardFormat: style.Format?.StandardFormat);

        AddedStyles.Add(addedStyle);

        var newId = StyleElements.Count;
        StyleElements.Add(new StyleElement(style, name, visibility));
        return newId;
    }

    private int? AssignDateTimeStyleId(in ImmutableStyle style)
    {
        if (style.Format is not null || _defaultDateTimeFormat is not { } dateTimeFormat)
            return null;

        // Optionally add another style for DateTime when there is no explicit number format in the new style.
        var dateTimeStyle = style with { Format = dateTimeFormat };

        if (_styleDictionary.TryGetValue(dateTimeStyle, out var existing))
            return existing.Id;

        var dateTimeId = AssignStyleId(dateTimeStyle, name: null, visibility: null);
        _styleDictionary[dateTimeStyle] = new StyleId(dateTimeId, dateTimeId);
        return dateTimeId;
    }

    public bool TryAddNamedStyle(string name, Style style, StyleNameVisibility? visibility,
        [NotNullWhen(true)] out StyleId? styleId)
    {
        styleId = null;

        var namedStyles = NamedStyles ??= new(StringComparer.OrdinalIgnoreCase);
        if (namedStyles.ContainsKey(name))
            return false;

        var immutableStyle = ImmutableStyle.From(style, DefaultFont);

        // When adding a named style, we don't want to check for an existing style for potential reuse.
        // The reason is that if the named style would change, we only want it to affect parts where the
        // style name has been used explicitly.
        // Style names only refer to the main style, not the DateTime style.
        // Since the DateTime style doesn't have a name, it can be reused like regular styles.
        var id = AssignStyleId(immutableStyle, name, visibility);
        var dateTimeId = AssignDateTimeStyleId(immutableStyle) ?? id;
        styleId = new StyleId(id, dateTimeId);
        namedStyles[name] = (styleId, namedStyles.Count);
        return true;
    }

    public StyleId? GetStyleIdOrDefault(string name)
    {
        return NamedStyles is { } namedStyles && namedStyles.TryGetValue(name, out var value)
            ? value.StyleId
            : null;
    }

    public List<(string, ImmutableStyle, AddedStyle, StyleNameVisibility)>? GetEmbeddedNamedStyles()
    {
        if (NamedStyles is null)
            return null;

        List<(string, ImmutableStyle, AddedStyle, StyleNameVisibility)>? result = null;

        var index = -1;
        foreach (var (style, name, visibility) in StyleElements)
        {
            index++;

            if (name is null) continue;
            if (visibility is not { } nameVisibility) continue;

            var addedStyle = AddedStyles[index];

            result ??= [];
            result.Add((name, style, addedStyle, nameVisibility));
        }

        return result;
    }
}
