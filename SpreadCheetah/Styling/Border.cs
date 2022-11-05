namespace SpreadCheetah.Styling;

/// <summary>
/// Represents all border parts of a <see cref="Style"/>.
/// </summary>
public sealed class Border : IEquatable<Border>
{
    /// <summary>Left border for the cell.</summary>
    public EdgeBorder Left { get; set; } = new();

    /// <summary>Right border for the cell.</summary>
    public EdgeBorder Right { get; set; } = new();

    /// <summary>Top border for the cell.</summary>
    public EdgeBorder Top { get; set; } = new();

    /// <summary>Bottom border for the cell.</summary>
    public EdgeBorder Bottom { get; set; } = new();

    /// <summary>Diagonal border for the cell.</summary>
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
