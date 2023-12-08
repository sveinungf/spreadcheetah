namespace SpreadCheetah.Images;

[Flags]
internal enum ImageCanvasOptions
{
    None,
    MoveWithCells = 1,
    ResizeWithCells = 2,
    Scaled = 4,
    Dimensions = 8,
    FillCell = 16
}
