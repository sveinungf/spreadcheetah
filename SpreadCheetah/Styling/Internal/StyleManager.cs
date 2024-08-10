using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.Styling.Internal;

internal sealed class StyleManager
{
    private readonly record struct StyleIdentifiers(
        int StyleId);

    private readonly record struct EmbeddedNamedStyle(
        StyleNameVisibility Visibility);

    private readonly Dictionary<ImmutableStyle, StyleIdentifiers> _styles = [];
    private readonly NumberFormat? _defaultDateTimeFormat;
    private Dictionary<string, StyleId>? _namedStyles;
    private Dictionary<string, EmbeddedNamedStyle>? _embeddedNamedStyles;

    public DefaultStyling? DefaultStyling { get; }

    public StyleManager(NumberFormat? defaultDateTimeFormat)
    {
        _defaultDateTimeFormat = defaultDateTimeFormat;

        // If we have any style, the built-in default style must be the first one (meaning the first <xf> element in styles.xml).
        var defaultFont = new ImmutableFont(null, false, false, false, Font.DefaultSize, null);
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
        if (_styles.TryGetValue(style, out var existingIdentifiers))
            return existingIdentifiers.StyleId;

        var id = _styles.Count;
        _styles[style] = new StyleIdentifiers(id);
        return id;
    }

    public bool TryAddNamedStyle(string name, Style style, StyleNameVisibility? nameVisibility,
        [NotNullWhen(true)] out StyleId? styleId)
    {
        styleId = null;

        var namedStyles = _namedStyles ??= new(StringComparer.OrdinalIgnoreCase);
        if (namedStyles.ContainsKey(name))
            return false;

        // TODO: When there is a default DateTime number format, two style IDs will be created.
        // TODO: How should this be handled for a named style?
        // TODO: Maybe the named style should only refer to the regular style, not the DateTime style.
        styleId = AddStyleIfNotExists(ImmutableStyle.From(style));
        namedStyles[name] = styleId;

        if (nameVisibility is { } visibility)
        {
            var embeddedNamedStyles = _embeddedNamedStyles ??= new(StringComparer.OrdinalIgnoreCase);
            embeddedNamedStyles[name] = new EmbeddedNamedStyle(visibility);
        }

        return true;
    }

    public StyleId? GetStyleIdOrDefault(string name) => _namedStyles?.GetValueOrDefault(name);

    public List<ImmutableStyle> GetOrderedStyles()
    {
        // The order of Dictionary.Keys is not guaranteed, so we make sure the styles are sorted by the StyleId here.
        return _styles.OrderBy(x => x.Value.StyleId).Select(x => x.Key).ToList();
    }

    public List<(string, ImmutableStyle, StyleNameVisibility)>? GetEmbeddedNamedStyles()
    {
        if (_embeddedNamedStyles is null || _namedStyles is null)
            return null;

        var result = new List<(string, ImmutableStyle, StyleNameVisibility)>(_embeddedNamedStyles.Count);
        var styleIdToStyle = _styles.ToDictionary(x => x.Value.StyleId, x => x.Key);

        foreach (var (name, embeddedNameStyle) in _embeddedNamedStyles)
        {
            if (!_namedStyles.TryGetValue(name, out var styleId))
                continue;
            if (!styleIdToStyle.TryGetValue(styleId.Id, out var style))
                continue;

            result.Add((name, style, embeddedNameStyle.Visibility));
        }

        return result;
    }
}
