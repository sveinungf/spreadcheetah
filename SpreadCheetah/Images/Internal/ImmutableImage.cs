namespace SpreadCheetah.Images.Internal;

internal readonly record struct ImmutableImage(
    int EmbeddedImageId,
    int ActualImageWidth,
    int ActualImageHeight,
    ImageAnchor Anchor,
    ImageSize? DesiredSize);