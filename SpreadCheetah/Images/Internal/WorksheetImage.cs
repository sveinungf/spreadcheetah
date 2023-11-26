namespace SpreadCheetah.Images.Internal;

internal readonly record struct WorksheetImage(
    ImageCanvas Canvas,
    EmbeddedImage EmbeddedImage,
    ImmutableImage Image, // TODO: Remove
    int ImageNumber);