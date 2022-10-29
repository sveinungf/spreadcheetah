namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableStyle(
    ImmutableBorder Border,
    ImmutableFill Fill,
    ImmutableFont Font,
    string? NumberFormat)
{
    public static ImmutableStyle From(Style style) => new(
        Border: ImmutableBorder.From(style.Border),
        Fill: ImmutableFill.From(style.Fill),
        Font: ImmutableFont.From(style.Font),
        NumberFormat: style.NumberFormat);
}
