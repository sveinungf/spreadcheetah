namespace SpreadCheetah.Styling;

[Flags]
internal enum DiagonalBorderType
{
    None = 0,
    DiagonalUp = 1,
    DiagonalDown = 2,
    CrissCross = DiagonalUp | DiagonalDown
}
