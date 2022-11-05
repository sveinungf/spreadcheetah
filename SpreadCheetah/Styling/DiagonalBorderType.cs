namespace SpreadCheetah.Styling;

/// <summary>
/// Values that specifies the type of diagonal border used on cells.
/// </summary>
[Flags]
public enum DiagonalBorderType
{
    /// <summary>No border.</summary>
    None = 0,

    /// <summary>Diagonal up.</summary>
    DiagonalUp = 1,

    /// <summary>Diagonal down.</summary>
    DiagonalDown = 2,

    /// <summary>Diagonal up and diagonal down.</summary>
    CrissCross = DiagonalUp | DiagonalDown
}
