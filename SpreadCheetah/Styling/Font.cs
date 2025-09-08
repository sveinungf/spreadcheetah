using SpreadCheetah.Helpers;
using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the font part of a <see cref="Style"/>.
/// </summary>
public sealed record Font
{
    /// <summary>Font name. Defaults to Calibri.</summary>
    public string? Name
    {
        get => _name;
        set => _name = Guard.FontNameLengthInRange(value);
    }

    private string? _name;

    /// <summary>Bold font weight. Defaults to <see langword="false"/>.</summary>
    public bool Bold { get; set; }

    /// <summary>Italic font type. Defaults to <see langword="false"/>.</summary>
    public bool Italic { get; set; }

    /// <summary>Adds a horizontal line through the center of the characters. Defaults to <see langword="false"/>.</summary>
    public bool Strikethrough { get; set; }

    /// <summary>Font underline. Defaults to no underline.</summary>
    public Underline Underline
    {
        get => _underline;
        set => _underline = Guard.DefinedEnumValue(value);
    }

    private Underline _underline;

    /// <summary>Font size. Defaults to 11.</summary>
    public double Size
    {
        get => ActualSize ?? DefaultFont.DefaultSize;
        set => ActualSize = Guard.FontSizeInRange(value);
    }

    /// <summary>
    /// Keep track of whether or not the user has set a value for the 'Size' property.
    /// We need to know when we should fall back to the size from DefaultFont.
    /// Since the 'Size' property defaults to 11, we need to know when 'Size' has explicitly been set to 11.
    /// </summary>
    internal double? ActualSize { get; set; }

    /// <summary>ARGB (alpha, red, green, blue) color of the font.</summary>
    public Color? Color { get; set; }
}
