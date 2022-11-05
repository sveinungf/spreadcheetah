using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableBorder(
    ImmutableEdgeBorder Left,
    ImmutableEdgeBorder Right,
    ImmutableEdgeBorder Top,
    ImmutableEdgeBorder Bottom,
    ImmutableDiagonalBorder Diagonal)
{
    public static ImmutableBorder From(Border border) => new(
        Left: ImmutableEdgeBorder.From(border.Left),
        Right: ImmutableEdgeBorder.From(border.Right),
        Top: ImmutableEdgeBorder.From(border.Top),
        Bottom: ImmutableEdgeBorder.From(border.Bottom),
        Diagonal: ImmutableDiagonalBorder.From(border.Diagonal));
}
