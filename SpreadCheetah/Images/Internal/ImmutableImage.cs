namespace SpreadCheetah.Images.Internal;

internal readonly record struct ImmutableImage(
    int EmbeddedImageId,
    string Name,
    int ActualImageWidth,
    int ActualImageHeight,
    ImageSize? DesiredSize);