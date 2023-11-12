namespace SpreadCheetah.Images;

public sealed record ImageBoundedCanvas : ImageCanvas
{
    internal int ToColumn { get; init; }
    internal int ToRow { get; init; }
    internal bool ResizeWithCellsValue { get; private init; }

    private ImageBoundedCanvas()
    {
    }

    internal ImageBoundedCanvas(ImageCanvas imageCanvas)
    {
        Column = imageCanvas.Column;
        Row = imageCanvas.Row;
        MoveWithCellsValue = imageCanvas.MoveWithCellsValue;
    }

    // TODO: Hmm, possible to reuse implementation from base class?
    public new ImageBoundedCanvas MoveWithCells(bool moveWithCells = true)
    {
        return this with { MoveWithCellsValue = moveWithCells };
    }

    public ImageBoundedCanvas ResizeWithCells(bool resizeWithCells = true)
    {
        return this with { ResizeWithCellsValue = resizeWithCells };
    }
}
