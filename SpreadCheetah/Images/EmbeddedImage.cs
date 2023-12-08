using System.Diagnostics;

namespace SpreadCheetah.Images;

/// <summary>
/// Reference for an embedded image. Images are embedded with <see cref="Spreadsheet.EmbedImageAsync"/>.
/// </summary>
public sealed class EmbeddedImage
{
    /// <summary>The width of the image in pixels.</summary>
    public int Width { get; }

    /// <summary>The height of the image in pixels.</summary>
    public int Height { get; }

    /// <summary>
    /// Should be globally unique in the spreadsheet. First image has ID = 1, second image has ID = 2, etc.
    /// </summary>
    internal int Id { get; }
    internal Guid SpreadsheetGuid { get; }

    internal EmbeddedImage(int width, int height, int id, Guid spreadsheetGuid)
    {
        Debug.Assert(id > 0);
        Width = width;
        Height = height;
        Id = id;
        SpreadsheetGuid = spreadsheetGuid;
    }
}
