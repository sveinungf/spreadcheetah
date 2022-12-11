using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the diagonal border part of a <see cref="Style"/>.
/// </summary>
public sealed record DiagonalBorder
{
    /// <summary>Border style. Defaults to None.</summary>
    public BorderStyle BorderStyle { get; set; } = BorderStyle.None;

    /// <summary>ARGB (alpha, red, green, blue) color of the border.</summary>
    public Color? Color { get; set; }

    /// <summary>Diagonal up, diagonal down, or both. Defaults to None.</summary>
    public DiagonalBorderType Type { get; set; } = DiagonalBorderType.None;
}