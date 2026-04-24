namespace SpreadCheetah.Styling.Internal;

internal readonly record struct AddedDifferentialStyle
{
    // TODO: Add more parts of the style.
    public required ImmutableFill Fill { get; init; }
}
