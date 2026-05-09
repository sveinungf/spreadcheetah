namespace SpreadCheetah.Styling.Internal;

internal readonly record struct AddedDifferentialStyle
{
    // TODO: Add more parts of the style.
    public required ImmutableFill Fill { get; init; }
    public required ImmutableFont Font { get; init; }
    public required string? Format { get; init; }
}
