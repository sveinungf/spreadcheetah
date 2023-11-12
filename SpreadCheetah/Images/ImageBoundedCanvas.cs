namespace SpreadCheetah.Images;

public readonly record struct ImageBoundedCanvas
{
    internal int FromColumn { get; private init; }
    internal int FromRow { get; private init; }
    internal int ToColumn { get; init; }
    internal int ToRow { get; init; }
    internal bool MoveWithCellsValue { get; private init; }
    internal bool ResizeWithCellsValue { get; private init; }

    public ImageBoundedCanvas()
    {
        FromColumn = 1;
        FromRow = 1;
        ToColumn = 2;
        ToRow = 2;
        MoveWithCellsValue = true;
        ResizeWithCellsValue = true;
    }

    internal ImageBoundedCanvas(ImageCanvas imageCanvas)
    {
        FromColumn = imageCanvas.Column;
        FromRow = imageCanvas.Row;
        MoveWithCellsValue = imageCanvas.MoveWithCellsValue;
    }

    public ImageBoundedCanvas MoveWithCells(bool moveWithCells = true)
    {
        return this with { MoveWithCellsValue = moveWithCells };
    }

    public ImageBoundedCanvas ResizeWithCells(bool resizeWithCells = true)
    {
        return this with { ResizeWithCellsValue = resizeWithCells };
    }
}
