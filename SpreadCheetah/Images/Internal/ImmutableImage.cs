namespace SpreadCheetah.Images.Internal;

internal readonly record struct ImmutableImage(
    int EmbeddedImageId,
    (int Width, int Height) OriginalDimensions,
    ImageAnchor Anchor,
    ImageSize? DesiredSize);