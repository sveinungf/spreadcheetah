namespace SpreadCheetah.Styling;

internal class Border
{
    public EdgeBorder Left { get; set; } = new();
    public EdgeBorder Right { get; set; } = new();
    public EdgeBorder Top { get; set; } = new();
    public EdgeBorder Bottom { get; set; } = new();
    public DiagonalBorder Diagonal { get; set; } = new();
}
