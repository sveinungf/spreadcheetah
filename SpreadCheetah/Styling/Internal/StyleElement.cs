namespace SpreadCheetah.Styling.Internal;

internal readonly record struct StyleElement(ImmutableStyle Style, string? Name, StyleNameVisibility? Visibility);
