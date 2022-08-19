namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableStyle(
    ImmutableFill Fill,
    ImmutableFont Font,
    string? NumberFormat)
{
    public static ImmutableStyle From(Style style) => new(
        Fill: ImmutableFill.From(style.Fill),
        Font: ImmutableFont.From(style.Font),
        NumberFormat: style.NumberFormat);
}
