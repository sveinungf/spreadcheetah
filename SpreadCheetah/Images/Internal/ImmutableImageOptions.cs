namespace SpreadCheetah.Images.Internal;

internal readonly record struct ImmutableImageOptions(
    string Name,
    ImageSize? Size);