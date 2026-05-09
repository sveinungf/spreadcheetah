using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the fill part of a <see cref="DifferentialStyle"/>.
/// </summary>
public sealed record DifferentialFill
{
    /// <summary>ARGB (alpha, red, green, blue) color of the fill.</summary>
    public Color? Color { get; set; }
}
