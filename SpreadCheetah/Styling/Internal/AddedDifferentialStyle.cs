namespace SpreadCheetah.Styling.Internal;

internal readonly record struct AddedDifferentialStyle
{
    public required ImmutableBorder Border { get; init; }
    public required ImmutableFill Fill { get; init; }
    public required ImmutableFont Font { get; init; }
    public required string? Format { get; init; }
}
