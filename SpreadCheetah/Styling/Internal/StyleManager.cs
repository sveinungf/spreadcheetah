namespace SpreadCheetah.Styling.Internal;

internal sealed class StyleManager
{
    private readonly record struct StyleIdentifiers(
        int StyleId);

    private readonly record struct EmbeddedNamedStyle(
        StyleNameVisibility Visibility);

    private readonly Dictionary<ImmutableStyle, StyleIdentifiers> _styles = [];
    private Dictionary<string, StyleId>? _namedStyles;
    private Dictionary<string, EmbeddedNamedStyle>? _embeddedNamedStyles;

    public DefaultStyling? DefaultStyling { get; private set; }

    public StyleId AddStyleIfNotExists(in ImmutableStyle style, NumberFormat? defaultDateTimeFormat)
    {
        var id = AddStyleIfNotExistsInternal(style);

        if (defaultDateTimeFormat is null || style.Format is not null)
            return new StyleId(id, id);

        // Optionally add another style for DateTime when there is no explicit number format in the new style.
        var dateTimeStyle = style with { Format = defaultDateTimeFormat };
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

    public bool StyleNameExists(string name)
    {
        return _namedStyles is { } namedStyles && namedStyles.ContainsKey(name);
    }

    public void AddNamedStyle(string name, StyleId styleId, StyleNameVisibility? visibility)
    {
        var namedStyles = _namedStyles ??= new(StringComparer.OrdinalIgnoreCase);
        namedStyles[name] = styleId;

        if (visibility is { } vis)
        {
            var embeddedNamedStyles = _embeddedNamedStyles ??= new(StringComparer.OrdinalIgnoreCase);
            embeddedNamedStyles[name] = new EmbeddedNamedStyle(vis);
        }
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

    public void AddDefaultStyle(NumberFormat? defaultDateTimeFormat)
    {
        var defaultFont = new ImmutableFont(null, false, false, false, Font.DefaultSize, null);
        var defaultStyle = new ImmutableStyle(new ImmutableAlignment(), new ImmutableBorder(), new ImmutableFill(), defaultFont, null);
        var styleId = AddStyleIfNotExists(defaultStyle, defaultDateTimeFormat);

        if (styleId.Id != styleId.DateTimeId)
            DefaultStyling = new DefaultStyling(styleId.DateTimeId);
    }
}
