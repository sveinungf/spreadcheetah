namespace SpreadCheetah.Images;

public sealed class ImageSize
{
    internal decimal? ScaleValue { get; private set; }
    internal (int Width, int Height)? DimensionsValue { get; private set; }

    private ImageSize()
    {
    }

    public static ImageSize Dimensions(int width, int height)
    {
        // TODO: Sanity check for arguments
        return new ImageSize { DimensionsValue = (width, height) };
    }

    public static ImageSize Scale(decimal scale)
    {
        // TODO: Sanity check for argument
        return new ImageSize { ScaleValue = scale };
    }

    // TODO: Fill cell option: A way of doing this could be to pass two cell references when adding images. Can fill multiple cells this way.

    // TODO: Fit cell option: Keep aspect ratio, but with maximum image size inside a cell (or multiple cells).
    // TODO: Should probably pass two cell references for this option as well, otherwise resizing with cells won't work.
    // TODO: Should have some alignment option as well? But that won't work with offsets.
    // TODO: Will need to know column widths and row heights. Column widths should be known. Total row height should be passed as argument? Note: Can be for multiple columns/rows.
    // TODO: Perhaps this option can come in a later version.

    // TODO: Different image options when adding for a single cell reference compared to for two cell references?
}