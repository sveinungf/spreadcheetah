using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableBorder
{
    public required ImmutableEdgeBorder? Left { get; init; }
    public required ImmutableEdgeBorder? Right { get; init; }
    public required ImmutableEdgeBorder? Top { get; init; }
    public required ImmutableEdgeBorder? Bottom { get; init; }
    public required ImmutableDiagonalBorder? Diagonal { get; init; }

    public static ImmutableBorder From(Border border) => new()
    {
        Left = ImmutableEdgeBorder.From(border.Left),
        Right = ImmutableEdgeBorder.From(border.Right),
        Top = ImmutableEdgeBorder.From(border.Top),
        Bottom = ImmutableEdgeBorder.From(border.Bottom),
        Diagonal = ImmutableDiagonalBorder.From(border.Diagonal)
    };

    public static ImmutableBorder Default => new()
    {
        Left = new ImmutableEdgeBorder(),
        Right = new ImmutableEdgeBorder(),
        Top = new ImmutableEdgeBorder(),
        Bottom = new ImmutableEdgeBorder(),
        Diagonal = new ImmutableDiagonalBorder()
    };
}
