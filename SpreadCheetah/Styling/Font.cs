using SpreadCheetah.Helpers;
using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the font part of a <see cref="Style"/>.
/// </summary>
public sealed record Font
{
    internal const double DefaultSize = 11;

    // TODO: Update comment
    /// <summary>Font name. Defaults to Calibri.</summary>
    public string? Name
    {
        get => _name;
        set => _name = Guard.MaxLength(value, 31);
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
        get => _size;
        set => _size = Guard.FontSizeInRange(value);
    }

    private double _size = DefaultSize;

    /// <summary>ARGB (alpha, red, green, blue) color of the font.</summary>
    public Color? Color { get; set; }
}
