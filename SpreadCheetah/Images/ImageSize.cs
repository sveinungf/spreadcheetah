namespace SpreadCheetah.Images;

public sealed class ImageSize
{
    internal decimal? ScaleValue { get; private set; }
    internal (int Width, int Height)? DimensionsValue { get; private set; }

    private ImageSize()
    {
    }

    public static ImageSize Scale(decimal scale)
    {
        return new ImageSize { ScaleValue = scale };
    }

    public static ImageSize Dimensions(int width, int height)
    {
        return new ImageSize { DimensionsValue = (width, height) };
    }

    public static ImageSize FillCell()
    {
        return new ImageSize();
    }

    public static ImageSize FitCell(int columnWidth, int rowHeight)
    {
        return new ImageSize();
    }
}