using SpreadCheetah.Helpers;
using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the font part of a <see cref="DifferentialStyle"/>.
/// </summary>
public sealed record DifferentialFont
{
    /// <summary>Bold font weight. Defaults to <see langword="false"/>.</summary>
    public bool Bold { get; set; }

    /// <summary>Italic font type. Defaults to <see langword="false"/>.</summary>
    public bool Italic { get; set; }

    /// <summary>Adds a horizontal line through the center of the characters. Defaults to <see langword="false"/>.</summary>
    public bool Strikethrough { get; set; }

    // TODO: Verify if all the underline options are supported.
    /// <summary>Font underline. Defaults to no underline.</summary>
    public Underline Underline { get; set => field = Guard.DefinedEnumValue(value); }

    /// <summary>ARGB (alpha, red, green, blue) color of the font.</summary>
    public Color? Color { get; set; }
}
