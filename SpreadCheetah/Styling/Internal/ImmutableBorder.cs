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
        Left = ImmutableEdgeBorder.From(border.Left),
        Right = ImmutableEdgeBorder.From(border.Right),
        Top = ImmutableEdgeBorder.From(border.Top),
        Bottom = ImmutableEdgeBorder.From(border.Bottom),
        Diagonal = ImmutableDiagonalBorder.From(border.Diagonal),
        IsConditionalFormatBorder = false
    };

    public static ImmutableBorder From(ConditionalFormatBorder border) => new()
    {
        Left = ImmutableEdgeBorder.From(border.Left),
        Right = ImmutableEdgeBorder.From(border.Right),
        Top = ImmutableEdgeBorder.From(border.Top),
        Bottom = ImmutableEdgeBorder.From(border.Bottom),
        Diagonal = null,
        IsConditionalFormatBorder = true
    };
}
