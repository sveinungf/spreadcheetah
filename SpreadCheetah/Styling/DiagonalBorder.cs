using System.Drawing;

namespace SpreadCheetah.Styling;

internal sealed class DiagonalBorder : IEquatable<DiagonalBorder>
{
    public BorderStyle BorderStyle { get; set; } = BorderStyle.None;
    public Color? Color { get; set; }
    public DiagonalBorderType Type { get; set; } = DiagonalBorderType.None;

    /// <inheritdoc/>
    public bool Equals(DiagonalBorder? other) => other != null
        && BorderStyle == other.BorderStyle
        && Type == other.Type
        && EqualityComparer<Color?>.Default.Equals(Color, other.Color);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is DiagonalBorder other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(BorderStyle, Color, Type);
}