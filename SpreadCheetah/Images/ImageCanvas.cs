using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using System.Runtime.InteropServices;

namespace SpreadCheetah.Images;

/// <summary>
/// Represents the placement and size of an image in a worksheet.
/// </summary>
[StructLayout(LayoutKind.Auto)]
public readonly record struct ImageCanvas
{
    internal ImageCanvasOptions Options { get; }
    /// <summary>Starts at 0, meaning 'A' = 0</summary>
    internal ushort FromColumn { get; }
    /// <summary>Starts at 0, meaning row 1 = 0</summary>
    internal uint FromRow { get; }
    /// <summary>Starts at 0, meaning 'A' = 0</summary>
    internal ushort ColumnCount { get; }
    /// <summary>Starts at 0, meaning row 1 = 0</summary>
    internal uint RowCount { get; }
    internal float ScaleValue { get; }
    internal ushort DimensionWidth { get; }
    internal ushort DimensionHeight { get; }

    private ImageCanvas(SingleCellRelativeReference fromReference, ImageCanvasOptions options, ushort columnCount = 0, uint rowCount = 0,
        float scaleValue = 0, ushort dimensionWidth = 0, ushort dimensionHeight = 0)
    {
        FromColumn = (ushort)(fromReference.Column - 1);
        FromRow = fromReference.Row - 1;
        Options = options;
        ColumnCount = columnCount;
        RowCount = rowCount;
        ScaleValue = scaleValue;
        DimensionWidth = dimensionWidth;
        DimensionHeight = dimensionHeight;
    }

    /// <summary>
    /// Show the image in its original size with its upper left corner at <paramref name="upperLeftReference"/>.
    /// The cell reference must be in the A1 reference style, e.g. "A1" or "C4".
    /// The <paramref name="moveWithCells"/> parameter decides whether or not the image should move when
    /// cells further up or left change size.
    /// </summary>
    public static ImageCanvas OriginalSize(ReadOnlySpan<char> upperLeftReference, bool moveWithCells = true)
    {
        var reference = SingleCellRelativeReference.Create(upperLeftReference);
        var options = moveWithCells ? ImageCanvasOptions.MoveWithCells : ImageCanvasOptions.None;
        return new ImageCanvas(reference, options);
    }

    /// <summary>
    /// Show the image with the specified dimensions with its upper left corner at <paramref name="upperLeftReference"/>.
    /// The cell reference must be in the A1 reference style, e.g. "A1" or "C4".
    /// The <paramref name="moveWithCells"/> parameter decides whether or not the image should move when
    /// cells further up or left change size.
    /// </summary>
    public static ImageCanvas Dimensions(ReadOnlySpan<char> upperLeftReference, int width, int height, bool moveWithCells = true)
    {
        var reference = SingleCellRelativeReference.Create(upperLeftReference);
        width.EnsureValidImageDimension();
        height.EnsureValidImageDimension();
        var options = moveWithCells ? ImageCanvasOptions.Dimensions | ImageCanvasOptions.MoveWithCells : ImageCanvasOptions.Dimensions;
        return new ImageCanvas(reference, options, dimensionWidth: (ushort)width, dimensionHeight: (ushort)height);
    }

    /// <summary>
    /// Show the image with the specified scale with its upper left corner at <paramref name="upperLeftReference"/>.
    /// The <paramref name="scale"/> must be between 0.0 and 1000. A <paramref name="scale"/> less than 1.0 will make the
    /// image appear smaller than its original size, while a <paramref name="scale"/> greater than 1.0 will make the
    /// image appear larger. The cell reference must be in the A1 reference style, e.g. "A1" or "C4".
    /// The <paramref name="moveWithCells"/> parameter decides whether or not the image should move when
    /// cells further up or left change size.
    /// </summary>
    public static ImageCanvas Scaled(ReadOnlySpan<char> upperLeftReference, float scale, bool moveWithCells = true)
    {
        var reference = SingleCellRelativeReference.Create(upperLeftReference);

        if (scale <= 0.0f)
            ThrowHelper.ImageScaleTooSmall(nameof(scale), scale);
        if (scale > 1000f)
            ThrowHelper.ImageScaleTooLarge(nameof(scale), scale);

        var options = moveWithCells ? ImageCanvasOptions.Scaled | ImageCanvasOptions.MoveWithCells : ImageCanvasOptions.Scaled;
        return new ImageCanvas(reference, options, scaleValue: scale);
    }

    private static ImageCanvas FillCell(SingleCellRelativeReference upperLeft, ushort columnCount, uint rowCount, bool moveWithCells, bool resizeWithCells)
    {
        var options = ImageCanvasOptions.FillCell;
        if (moveWithCells)
            options |= ImageCanvasOptions.MoveWithCells;
        if (resizeWithCells)
            options |= ImageCanvasOptions.ResizeWithCells;

        return new ImageCanvas(upperLeft, options, columnCount, rowCount);
    }

    /// <summary>
    /// Show the image so that it fills the cell at <paramref name="cellReference"/>.
    /// The cell reference must be in the A1 reference style, e.g. "A1" or "C4".
    /// The <paramref name="moveWithCells"/> parameter decides whether or not the image should move when
    /// cells further up or left change size.
    /// The <paramref name="resizeWithCells"/> parameter decides whether or not the image should resize when
    /// the cell is being resized. <paramref name="resizeWithCells"/> can only be set to <see langword="true"/>
    /// if <paramref name="moveWithCells"/> is also set to <see langword="true"/>.
    /// </summary>
    public static ImageCanvas FillCell(ReadOnlySpan<char> cellReference, bool moveWithCells = true, bool resizeWithCells = true)
    {
        var upperLeft = SingleCellRelativeReference.Create(cellReference);
        if (resizeWithCells && !moveWithCells)
            ThrowHelper.ResizeAndMoveCellsCombinationNotSupported(nameof(resizeWithCells), nameof(moveWithCells));

        return FillCell(upperLeft, columnCount: 1, rowCount: 1, moveWithCells, resizeWithCells);
    }

    /// <summary>
    /// Show the image so that it fills the cell range from <paramref name="upperLeftReference"/> to <paramref name="lowerRightReference"/>.
    /// The cell references must be in the A1 reference style, e.g. "A1" or "C4".
    /// The <paramref name="moveWithCells"/> parameter decides whether or not the image should move when
    /// cells further up or left change size.
    /// The <paramref name="resizeWithCells"/> parameter decides whether or not the image should resize when
    /// the cell range is being resized. <paramref name="resizeWithCells"/> can only be set to <see langword="true"/>
    /// if <paramref name="moveWithCells"/> is also set to <see langword="true"/>.
    /// </summary>
    public static ImageCanvas FillCells(ReadOnlySpan<char> upperLeftReference, ReadOnlySpan<char> lowerRightReference, bool moveWithCells = true, bool resizeWithCells = true)
    {
        var upperLeft = SingleCellRelativeReference.Create(upperLeftReference);
        var lowerRight = SingleCellRelativeReference.Create(lowerRightReference);
        var columnCount = lowerRight.Column - upperLeft.Column;
        var rowCount = lowerRight.Row - upperLeft.Row;
        if (columnCount <= 0 || rowCount <= 0)
            ThrowHelper.FillCellRangeMustContainAtLeastOneCell(nameof(lowerRightReference));
        if (resizeWithCells && !moveWithCells)
            ThrowHelper.ResizeAndMoveCellsCombinationNotSupported(nameof(resizeWithCells), nameof(moveWithCells));

        return FillCell(upperLeft, (ushort)columnCount, rowCount, moveWithCells, resizeWithCells);
    }
}
