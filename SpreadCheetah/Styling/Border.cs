namespace SpreadCheetah.Styling;

/// <summary>
/// Represents all border parts of a <see cref="Style"/>.
/// </summary>
public sealed record Border
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
}
