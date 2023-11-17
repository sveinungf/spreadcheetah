namespace SpreadCheetah.Images;

[Flags]
internal enum ImageCanvasOptions
{
    None,
    MoveWithCells = 1,
    ResizeWithCells = 2,
    Scale = 4,
    Dimensions = 8,
    FillCell = 16
}
