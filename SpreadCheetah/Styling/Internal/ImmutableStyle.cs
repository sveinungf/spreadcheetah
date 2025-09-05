using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableStyle(
    ImmutableAlignment Alignment,
    ImmutableBorder Border,
    ImmutableFill Fill,
    ImmutableFont Font,
    NumberFormat? Format)
{
    public static ImmutableStyle From(Style style, DefaultFont? defaultFont) => new(
        Alignment: ImmutableAlignment.From(style.Alignment),
        Border: ImmutableBorder.From(style.Border),
        Fill: ImmutableFill.From(style.Fill),
        Font: ImmutableFont.From(style.Font, defaultFont),
        Format: style.Format);
}
