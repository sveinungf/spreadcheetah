namespace SpreadCheetah.Images.Internal;

internal readonly record struct ImmutableImageOptions(
    string Name,
    int ActualImageHeight,
    int ActualImageWidth,
    ImageSize? DesiredSize);