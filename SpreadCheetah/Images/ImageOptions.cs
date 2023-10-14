namespace SpreadCheetah.Images;

public sealed class ImageOptions
{
    // TODO: Is this mandatory in the XML? If so, generate a name.
    // TODO: Any limit on length?
    public string? Name { get; set; } 
    public ImageSize? Size { get; set; }

    // TODO: Offsets
}
