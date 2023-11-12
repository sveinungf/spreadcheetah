using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.Images;

public readonly record struct ImageCanvas
{
    internal int Column { get; private init; }
    internal int Row { get; private init; }
    internal bool MoveWithCellsValue { get; private init; }
    internal decimal? ScaleValue { get; private init; }
    internal (int Width, int Height)? DimensionsValue { get; private init; }

    public ImageCanvas()
    {
        Column = 1;
        Row = 1;
    }

    public static ImageCanvas From(string upperLeftReference)
    {
        var reference = SingleCellRelativeReference.Create(upperLeftReference);
        return new ImageCanvas
        {
            Column = reference.Column,
            Row = reference.Row
        };
    }

    public ImageCanvas Dimensions(int width, int height)
    {
        width.EnsureValidImageDimension(nameof(width));
        height.EnsureValidImageDimension(nameof(height));
        return this with { DimensionsValue = (width, height) };
    }

    public ImageCanvas Scale(decimal scale)
    {
        if (scale <= 0.0m)
            ThrowHelper.ImageScaleTooSmall(nameof(scale), scale);
        if (scale > 1000m)
            ThrowHelper.ImageScaleTooLarge(nameof(scale), scale);

        return this with { ScaleValue = scale };
    }

    public ImageBoundedCanvas FillCell()
    {
        // TODO: Throw if dimensions or scale has been set
        return new ImageBoundedCanvas(this)
        {
            ToColumn = Column + 1,
            ToRow = Row + 1
        };
    }

    public ImageBoundedCanvas FillCellsTo(string lowerRightReference)
    {
        // TODO: Throw if dimensions or scale has been set
        var reference = SingleCellRelativeReference.Create(lowerRightReference);
        // TODO: Validate that reference is at least +1 compared to upperLeftReference
        return new ImageBoundedCanvas(this)
        {
            ToColumn = reference.Column,
            ToRow = reference.Row
        };
    }

    public ImageCanvas MoveWithCells(bool moveWithCells = true)
    {
        return this with { MoveWithCellsValue = moveWithCells };
    }
}
