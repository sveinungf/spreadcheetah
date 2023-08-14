namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableStyle(
    ImmutableAlignment Alignment,
    ImmutableBorder Border,
    ImmutableFill Fill,
    ImmutableFont Font,
    NumberFormat? Format)
{
    public static ImmutableStyle From(Style style) => new(
        Alignment: ImmutableAlignment.From(style.Alignment),
        Border: ImmutableBorder.From(style.Border),
        Fill: ImmutableFill.From(style.Fill),
        Font: ImmutableFont.From(style.Font),
        Format: style.Format);
}
