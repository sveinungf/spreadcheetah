using SpreadCheetah.Helpers;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.Images;

public sealed class EmbeddedImage
{
    public int Width { get; }
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

    internal void EnsureValidOptions(ImageOptions options, [CallerArgumentExpression(nameof(options))] string? paramName = null)
    {
        var originalDimensions = (Width, Height);
        if (options.Size?.ScaleValue is { } scale)
        {
            var (width, height) = originalDimensions.Scale(scale);
            width.EnsureValidImageDimension(paramName);
            height.EnsureValidImageDimension(paramName);
        }
    }
}
