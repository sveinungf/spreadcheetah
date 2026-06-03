using SpreadCheetah.Helpers;

namespace SpreadCheetah.ConditionalFormatting;

public sealed record ConditionalFormatStyle
{
    /// <summary>Border for the cell.</summary>
    public ConditionalFormatBorder Border
    {
        get => _border ??= new();
        set => _border = value;
    }

    private ConditionalFormatBorder? _border;
    internal ConditionalFormatBorder? GetBorderOrDefault() => _border;

    /// <summary>Fill for the cell.</summary>
    public ConditionalFormatFill Fill
    {
        get => _fill ??= new();
        set => _fill = value;
    }

    private ConditionalFormatFill? _fill;
    internal ConditionalFormatFill? GetFillOrDefault() => _fill;

    /// <summary>Font for the cell's value.</summary>
    public ConditionalFormatFont Font
    {
        get => _font ??= new();
        set => _font = value;
    }

    private ConditionalFormatFont? _font;
    internal ConditionalFormatFont? GetFontOrDefault() => _font;

    /// <summary>Format that defines how a number or <see cref="DateTime"/> cell should be displayed.</summary>
    public string? Format { get; set => field = Guard.MaxLength(value, 255); }

    internal bool IsDefault => this is
    {
        _border: null or { IsDefault: true },
        _fill: null or { IsDefault: true },
        _font: null or { IsDefault: true },
        Format: null
    };
}
