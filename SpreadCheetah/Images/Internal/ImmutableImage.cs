namespace SpreadCheetah.Images.Internal;

internal readonly record struct ImmutableImage(
    (int Width, int Height) OriginalDimensions,
    ImageAnchor Anchor,
    ImageOffset? Offset);