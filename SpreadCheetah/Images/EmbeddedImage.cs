namespace SpreadCheetah.Images;

public sealed class EmbeddedImage
{
    public int Height { get; }
    public int Width { get; }

    internal EmbeddedImage(int height, int width)
    {
        Height = height;
        Width = width;
    }
}
