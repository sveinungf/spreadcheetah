using SpreadCheetah.Images.Internal;

namespace SpreadCheetah.Images;

public sealed class ImageOptions
{
    // TODO: Remove
    public ImageSize? Size { get; set; }
    public ImageOffset? Offset { get; set; }

    public bool MoveWithCells { get; set; } = true;

    // TODO: In XML doc: Can't be set to true if MoveWithCells is false.
    // TODO: Not relevant when using Size = null, Size.Dimensions, and Size.Scale?
    public bool ResizeWithCells { get; set; }

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
