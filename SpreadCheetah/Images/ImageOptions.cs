namespace SpreadCheetah.Images;

/// <summary>
/// Provides options to be used when adding an image with <see cref="Spreadsheet.AddImage"/>.
/// </summary>
public sealed class ImageOptions
{
    /// <summary>
    /// Offset for the edges of the image relative to what is set by <see cref="ImageCanvas"/>.
    /// </summary>
    public ImageOffset? Offset { get; set; }
}
