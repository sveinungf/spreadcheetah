namespace SpreadCheetah.Styling;

internal sealed class Border : IEquatable<Border>
{
    public EdgeBorder Left { get; set; } = new();
    public EdgeBorder Right { get; set; } = new();
    public EdgeBorder Top { get; set; } = new();
    public EdgeBorder Bottom { get; set; } = new();
    public DiagonalBorder Diagonal { get; set; } = new();

    /// <inheritdoc/>
    public bool Equals(Border? other) => other != null
        && EqualityComparer<EdgeBorder>.Default.Equals(Left, other.Left)
        && EqualityComparer<EdgeBorder>.Default.Equals(Right, other.Right)
        && EqualityComparer<EdgeBorder>.Default.Equals(Top, other.Top)
        && EqualityComparer<EdgeBorder>.Default.Equals(Bottom, other.Bottom)
        && EqualityComparer<DiagonalBorder>.Default.Equals(Diagonal, other.Diagonal);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is Border other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(Left, Right, Top, Bottom, Diagonal);
}
