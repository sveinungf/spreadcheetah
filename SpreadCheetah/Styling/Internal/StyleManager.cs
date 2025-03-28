using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.Styling.Internal;

internal sealed class StyleManager
{
    private readonly NumberFormat? _defaultDateTimeFormat;
    private readonly Dictionary<ImmutableStyle, int> _styleDictionary = [];

    public DefaultStyling? DefaultStyling { get; }
    public List<StyleElement> StyleElements { get; } = [];
    public Dictionary<string, (StyleId StyleId, int NamedStyleIndex)>? NamedStyles { get; private set; }

    public StyleManager(NumberFormat? defaultDateTimeFormat)
    {
        _defaultDateTimeFormat = defaultDateTimeFormat;

        // If we have any style, the built-in default style must be the first one (meaning the first <xf> element in styles.xml).
        var defaultFont = new ImmutableFont() with { Size = Font.DefaultSize };
        var defaultStyle = new ImmutableStyle(new ImmutableAlignment(), new ImmutableBorder(), new ImmutableFill(), defaultFont, null);
        var styleId = AddStyleIfNotExists(defaultStyle);

        if (styleId.Id != styleId.DateTimeId)
            DefaultStyling = new DefaultStyling(styleId.DateTimeId);
    }

    public StyleId AddStyleIfNotExists(in ImmutableStyle style)
    {
        var id = AddStyleIfNotExistsInternal(style);

        if (_defaultDateTimeFormat is null || style.Format is not null)
            return new StyleId(id, id);

        // Optionally add another style for DateTime when there is no explicit number format in the new style.
        var dateTimeStyle = style with { Format = _defaultDateTimeFormat };
        var dateTimeId = AddStyleIfNotExistsInternal(dateTimeStyle);

        return new StyleId(id, dateTimeId);
    }

    private int AddStyleIfNotExistsInternal(in ImmutableStyle style)
    {
        if (_styleDictionary.TryGetValue(style, out var id))
            return id;

        var newId = StyleElements.Count;
        StyleElements.Add(new StyleElement(style, null, null));
        _styleDictionary[style] = newId;
        return newId;
    }

    public bool TryAddNamedStyle(string name, Style style, StyleNameVisibility? visibility,
        [NotNullWhen(true)] out StyleId? styleId)
    {
        styleId = null;

        var namedStyles = NamedStyles ??= new(StringComparer.OrdinalIgnoreCase);
        if (namedStyles.ContainsKey(name))
            return false;

        var immutableStyle = ImmutableStyle.From(style);

        var id = StyleElements.Count;
        StyleElements.Add(new StyleElement(immutableStyle, name, visibility));
        var dateTimeId = id;

        if (_defaultDateTimeFormat is not null && style.Format is null)
        {
            var dateTimeStyle = immutableStyle with { Format = _defaultDateTimeFormat };

            // Style names only refer to the regular style, not the DateTime style.
            // Since the DateTime style doesn't have a name, it can be reused like regular styles.
            dateTimeId = AddStyleIfNotExistsInternal(dateTimeStyle);
        }

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

    public List<(string, ImmutableStyle, StyleNameVisibility)>? GetEmbeddedNamedStyles()
    {
        if (NamedStyles is null)
            return null;

        List<(string, ImmutableStyle, StyleNameVisibility)>? result = null;

        foreach (var (style, name, visibility) in StyleElements)
        {
            if (name is null) continue;
            if (visibility is not { } nameVisibility) continue;

            result ??= [];
            result.Add((name, style, nameVisibility));
        }

        return result;
    }
}
