using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.Images;

public sealed class ImageSize
{
    private static ImageSize FillCellInstance { get; } = new() { FillCellValue = true };

    internal decimal? ScaleValue { get; private set; }
    internal (int Width, int Height)? DimensionsValue { get; private set; }
    internal bool FillCellValue { get; private set; }
    internal SingleCellRelativeReference? FillCellRangeLowerRightReference { get; private set; }

    private ImageSize()
    {
    }

    public static ImageSize Dimensions(int width, int height)
    {
        width.EnsureValidImageDimension(nameof(width));
        height.EnsureValidImageDimension(nameof(height));
        return new ImageSize { DimensionsValue = (width, height) };
    }

    public static ImageSize Scale(decimal scale)
    {
        if (scale <= 0.0m)
            ThrowHelper.ImageScaleTooSmall(nameof(scale), scale);
        if (scale > 1000m)
            ThrowHelper.ImageScaleTooLarge(nameof(scale), scale);

        return new ImageSize { ScaleValue = scale };
    }

    // TODO: Fill cell option: A way of doing this could be to pass two cell references when adding images. Can fill multiple cells this way.
    public static ImageSize FillCell() => FillCellInstance;

    public static ImageSize FillCellRange(string lowerRightCellReference)
    {
        return new ImageSize
        {
            FillCellValue = true,
            FillCellRangeLowerRightReference = SingleCellRelativeReference.Create(lowerRightCellReference)
        };
    }

    // TODO: Different image options when adding for a single cell reference compared to for two cell references?
}