using System.Diagnostics;

namespace SpreadCheetah.Images;

public sealed class EmbeddedImage
{
    public int Width { get; }
    public int Height { get; }

    /// <summary>
    /// Should be globally unique in the spreadsheet. First image has ID = 1, second image has ID = 2, etc.
    /// </summary>
    internal int Id { get; }

    internal EmbeddedImage(int width, int height, int id)
    {
        Debug.Assert(id > 0);
        Width = width;
        Height = height;
        Id = id;
    }
}
