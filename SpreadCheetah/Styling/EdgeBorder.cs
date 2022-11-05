using System.Drawing;

namespace SpreadCheetah.Styling;

internal sealed class EdgeBorder : IEquatable<EdgeBorder>
{
    public BorderStyle BorderStyle { get; set; } = BorderStyle.None;
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
