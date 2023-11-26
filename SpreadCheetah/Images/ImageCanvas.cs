using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.Images;

public readonly struct ImageCanvas
{
    internal ImageCanvasOptions Options { get; }
    /// <summary>Starts at 0, meaning 'A' = 0</summary>
    internal ushort FromColumn { get; }
    /// <summary>Starts at 0, meaning row 1 = 0</summary>
    internal uint FromRow { get; }
    /// <summary>Starts at 0, meaning 'A' = 0</summary>
    internal ushort ToColumn { get; }
    /// <summary>Starts at 0, meaning row 1 = 0</summary>
    internal uint ToRow { get; }
    internal float ScaleValue { get; }
    internal ushort DimensionWidth { get; }
    internal ushort DimensionHeight { get; }

    // TODO: Verify that default constructor places the image with original dimensions in A1

    private ImageCanvas(SingleCellRelativeReference fromReference, ImageCanvasOptions options, ushort toColumn = 0, uint toRow = 0, float scaleValue = 0, ushort dimensionWidth = 0, ushort dimensionHeight = 0)
    {
        FromColumn = (ushort)(fromReference.Column - 1);
        FromRow = (uint)(fromReference.Row - 1);
        Options = options;
        ToColumn = toColumn;
        ToRow = toRow;
        ScaleValue = scaleValue;
        DimensionWidth = dimensionWidth;
        DimensionHeight = dimensionHeight;
    }

    public static ImageCanvas OriginalSize(ReadOnlySpan<char> upperLeftReference, bool moveWithCells = true)
    {
        var reference = SingleCellRelativeReference.Create(upperLeftReference);
        var options = moveWithCells ? ImageCanvasOptions.MoveWithCells : ImageCanvasOptions.None;
        return new ImageCanvas(reference, options);
    }

    public static ImageCanvas Dimensions(ReadOnlySpan<char> upperLeftReference, int width, int height, bool moveWithCells = true)
    {
        // TODO: SingleCellRelativeReference.Column can be ushort?
        // TODO: SingleCellRelativeReference.Row can be uint?
        var reference = SingleCellRelativeReference.Create(upperLeftReference);
        width.EnsureValidImageDimension(nameof(width)); // TODO: Return ushort?
        height.EnsureValidImageDimension(nameof(height)); // TODO: Return ushort?
        var options = moveWithCells ? ImageCanvasOptions.Dimensions | ImageCanvasOptions.MoveWithCells : ImageCanvasOptions.Dimensions;
        return new ImageCanvas(reference, options, dimensionWidth: (ushort)width, dimensionHeight: (ushort)height);
    }

    // TODO: scale should be float?
    public static ImageCanvas Scaled(ReadOnlySpan<char> upperLeftReference, decimal scale, bool moveWithCells = true)
    {
        var reference = SingleCellRelativeReference.Create(upperLeftReference);

        if (scale <= 0.0m)
            ThrowHelper.ImageScaleTooSmall(nameof(scale), scale);
        if (scale > 1000m)
            ThrowHelper.ImageScaleTooLarge(nameof(scale), scale);

        var options = moveWithCells ? ImageCanvasOptions.Scaled | ImageCanvasOptions.MoveWithCells : ImageCanvasOptions.Scaled;
        return new ImageCanvas(reference, options, scaleValue: (float)scale);
    }

    private static ImageCanvas FillCell(SingleCellRelativeReference upperLeft, ushort lowerRightColumn, uint lowerRightRow, bool moveWithCells, bool resizeWithCells)
    {
        var options = ImageCanvasOptions.FillCell;
        if (moveWithCells)
            options |= ImageCanvasOptions.MoveWithCells;
        if (resizeWithCells)
            options |= ImageCanvasOptions.ResizeWithCells;

        return new ImageCanvas(upperLeft, options, lowerRightColumn, lowerRightRow);
    }

    // TODO: In XML doc: resizeWithCells can't be set to true if moveWithCells is false.
    public static ImageCanvas FillCell(ReadOnlySpan<char> cellReference, bool moveWithCells = true, bool resizeWithCells = true)
    {
        var upperLeft = SingleCellRelativeReference.Create(cellReference);
        if (resizeWithCells && !moveWithCells)
            ThrowHelper.ResizeAndMoveCellsCombinationNotSupported(nameof(resizeWithCells), nameof(moveWithCells));

        return FillCell(upperLeft, (ushort)upperLeft.Column, (uint)upperLeft.Row, moveWithCells, resizeWithCells);
    }

    // TODO: In XML doc: resizeWithCells can't be set to true if moveWithCells is false.
    public static ImageCanvas FillCells(ReadOnlySpan<char> upperLeftReference, ReadOnlySpan<char> lowerRightReference, bool moveWithCells = true, bool resizeWithCells = true)
    {
        var upperLeft = SingleCellRelativeReference.Create(upperLeftReference);
        var lowerRight = SingleCellRelativeReference.Create(lowerRightReference);
        if (lowerRight.Column <= upperLeft.Column || lowerRight.Row <= upperLeft.Row)
            ThrowHelper.FillCellRangeMustContainAtLeastOneCell(nameof(lowerRightReference));
        if (resizeWithCells && !moveWithCells)
            ThrowHelper.ResizeAndMoveCellsCombinationNotSupported(nameof(resizeWithCells), nameof(moveWithCells));

        return FillCell(upperLeft, (ushort)(lowerRight.Column - 1), (uint)(lowerRight.Row - 1), moveWithCells, resizeWithCells);
    }
}
