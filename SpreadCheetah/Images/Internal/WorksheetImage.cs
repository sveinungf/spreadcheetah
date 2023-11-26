namespace SpreadCheetah.Images.Internal;

internal readonly record struct WorksheetImage(
    ImageCanvas Canvas,
    EmbeddedImage EmbeddedImage,
    ImageOffset? Offset,
    int ImageNumber);