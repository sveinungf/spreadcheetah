using SpreadCheetah.ConditionalFormatting;
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
    public required bool IsConditionalFormatBorder { get; init; }

    public bool IsDefault => this is
    {
        Left.IsDefault: true,
        Right.IsDefault: true,
        Top.IsDefault: true,
        Bottom.IsDefault: true,
        Diagonal.IsDefault: true,
        IsConditionalFormatBorder: false
    };

    public static ImmutableBorder From(Border border) => new()
    {
        Left = ImmutableEdgeBorder.From(border.GetLeftOrDefault()),
        Right = ImmutableEdgeBorder.From(border.GetRightOrDefault()),
        Top = ImmutableEdgeBorder.From(border.GetTopOrDefault()),
        Bottom = ImmutableEdgeBorder.From(border.GetBottomOrDefault()),
        Diagonal = ImmutableDiagonalBorder.From(border.GetDiagonalOrDefault()),
        IsConditionalFormatBorder = false
    };

    public static ImmutableBorder From(ConditionalFormatBorder? border) => new()
    {
        Left = ImmutableEdgeBorder.From(border?.GetLeftOrDefault()),
        Right = ImmutableEdgeBorder.From(border?.GetRightOrDefault()),
        Top = ImmutableEdgeBorder.From(border?.GetTopOrDefault()),
        Bottom = ImmutableEdgeBorder.From(border?.GetBottomOrDefault()),
        Diagonal = null,
        IsConditionalFormatBorder = true
    };
}
