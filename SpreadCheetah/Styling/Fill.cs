using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the fill part of a <see cref="Style"/>.
/// </summary>
public sealed record Fill
{
    /// <summary>ARGB (alpha, red, green, blue) color of the fill.</summary>
    public Color? Color { get; set; }
}
