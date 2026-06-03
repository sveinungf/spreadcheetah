using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents style for one or more worksheet cells.
/// </summary>
public sealed record Style
{
    /// <summary>Alignment for the cell.</summary>
    public Alignment Alignment
    {
        get => _alignment ??= new();
        set => _alignment = value;
    }

    private Alignment? _alignment;
    internal Alignment? GetAlignmentOrDefault() => _alignment;

    /// <summary>Border for the cell.</summary>
    public Border Border
    {
        get => _border ??= new();
        set => _border = value;
    }

    private Border? _border;
    internal Border? GetBorderOrDefault() => _border;

    /// <summary>Fill for the cell.</summary>
    public Fill Fill
    {
        get => _fill ??= new();
        set => _fill = value;
    }

    private Fill? _fill;
    internal Fill? GetFillOrDefault() => _fill;

    /// <summary>Font for the cell's value.</summary>
    public Font Font
    {
        get => _font ??= new();
        set => _font = value;
    }

    private Font? _font;
    internal Font? GetFontOrDefault() => _font;

    /// <summary>Format that defines how a number or <see cref="DateTime"/> cell should be displayed.</summary>
    [Obsolete($"Use {nameof(Style)}.{nameof(Format)} instead")]
    public string? NumberFormat
    {
        get => Format?.CustomFormat;
        set => Format = value == null ? null : Styling.NumberFormat.FromLegacyString(value);
    }

    /// <summary>Format that defines how a number or <see cref="DateTime"/> cell should be displayed.</summary>
    public NumberFormat? Format { get; set; }

    /// <summary>
    /// Creates a style with the font color and underline formatting that Excel uses for HYPERLINK formulas.
    /// </summary>
    public static Style Hyperlink => new() { Font = { Color = Color.FromArgb(0x00467886), Underline = Underline.Single } };
}
