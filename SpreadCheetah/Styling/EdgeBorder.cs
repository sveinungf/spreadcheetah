using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents an edge border part of a <see cref="Style"/>.
/// </summary>
public sealed record EdgeBorder
{
    /// <summary>Border style. Defaults to None.</summary>
    public BorderStyle BorderStyle { get; set; } = BorderStyle.None;

    /// <summary>ARGB (alpha, red, green, blue) color of the border.</summary>
    public Color? Color { get; set; }
}
