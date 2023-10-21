namespace SpreadCheetah.Images;

public sealed class EmbeddedImage
{
    public int Width { get; }
    public int Height { get; }
    internal int Id { get; }

    internal EmbeddedImage(int width, int height, int id)
    {
        Width = width;
        Height = height;
        Id = id;
    }
}
