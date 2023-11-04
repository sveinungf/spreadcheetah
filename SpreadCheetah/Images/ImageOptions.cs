using SpreadCheetah.Images.Internal;

namespace SpreadCheetah.Images;

public sealed class ImageOptions
{
    public ImageSize? Size { get; set; }

    public bool MoveWithCells { get; set; } = true;

    // TODO: In XML doc: Can't be set to true if MoveWithCells is false.
    public bool ResizeWithCells { get; set; }

    // TODO: Offsets. Future version?

    // TODO: Use "Picture" instead of "Image" across the board, since that is what Excel is doing
    // TODO: For MoveAndSizeWithCells to work, the bottom right corner must reference a cell

    internal ImageAnchor GetAnchor()
    {
        return (MoveWithCells, ResizeWithCells) switch
        {
            (true, true) => ImageAnchor.TwoCell,
            (true, false) => ImageAnchor.OneCell,
            (false, false) => ImageAnchor.Absolute,
            _ => ImageAnchor.None
        };
    }
}
