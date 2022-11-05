using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents an edge border part of a <see cref="Style"/>.
/// </summary>
public sealed class EdgeBorder : IEquatable<EdgeBorder>
{
    /// <summary>Border style. Defaults to None.</summary>
    public BorderStyle BorderStyle { get; set; } = BorderStyle.None;

    /// <summary>ARGB (alpha, red, green, blue) color of the border.</summary>
    public Color? Color { get; set; }

    /// <inheritdoc/>
    public bool Equals(EdgeBorder? other) => other != null
        && BorderStyle == other.BorderStyle
        && EqualityComparer<Color?>.Default.Equals(Color, other.Color);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is EdgeBorder other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(BorderStyle, Color);
}
