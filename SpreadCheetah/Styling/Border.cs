namespace SpreadCheetah.Styling;

/// <summary>
/// Represents all border parts of a <see cref="Style"/>.
/// </summary>
public sealed record Border
{
    /// <summary>Left border for the cell.</summary>
    public EdgeBorder Left
    {
        get => _left ??= new();
        set => _left = value;
    }

    private EdgeBorder? _left;
    internal EdgeBorder? GetLeftOrDefault() => _left;

    /// <summary>Right border for the cell.</summary>
    public EdgeBorder Right
    {
        get => _right ??= new();
        set => _right = value;
    }

    private EdgeBorder? _right;
    internal EdgeBorder? GetRightOrDefault() => _right;

    /// <summary>Top border for the cell.</summary>
    public EdgeBorder Top
    {
        get => _top ??= new();
        set => _top = value;
    }

    private EdgeBorder? _top;
    internal EdgeBorder? GetTopOrDefault() => _top;

    /// <summary>Bottom border for the cell.</summary>
    public EdgeBorder Bottom
    {
        get => _bottom ??= new();
        set => _bottom = value;
    }

    private EdgeBorder? _bottom;
    internal EdgeBorder? GetBottomOrDefault() => _bottom;

    /// <summary>Diagonal border for the cell.</summary>
    public DiagonalBorder Diagonal
    {
        get => _diagonal ??= new();
        set => _diagonal = value;
    }

    private DiagonalBorder? _diagonal;
    internal DiagonalBorder? GetDiagonalOrDefault() => _diagonal;
}
