namespace SpreadCheetah.Styling;

[Flags]
public enum DiagonalBorderType
{
    None = 0,
    DiagonalUp = 1,
    DiagonalDown = 2,
    CrissCross = DiagonalUp | DiagonalDown
}
