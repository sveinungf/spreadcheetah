namespace SpreadCheetah.Images;

public sealed class ImageSize
{
    internal decimal? ScaleValue { get; private set; }
    internal (int Height, int Width)? DimensionsValue { get; private set; }

    private ImageSize()
    {
    }

    public static ImageSize Scale(decimal scale)
    {
        return new ImageSize { ScaleValue = scale };
    }

    public static ImageSize Dimensions(int height, int width)
    {
        return new ImageSize { DimensionsValue = (height, width) };
    }

    public static ImageSize FitCell()
    {
        return new ImageSize();
    }

    // TODO: Fit cell height option?
    // TODO: Fit cell width option?
}