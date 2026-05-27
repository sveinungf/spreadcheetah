using System.Drawing;
using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableDiagonalBorder
{
    public required BorderStyle BorderStyle { get; init; }
    public required Color? Color { get; init; }
    public required DiagonalBorderType Type { get; init; }

    public bool IsDefault => this is
    {
        BorderStyle: BorderStyle.None,
        Color: null,
        Type: DiagonalBorderType.None
    };

    public static ImmutableDiagonalBorder From(DiagonalBorder border) => new()
    {
        BorderStyle = border.BorderStyle,
        Color = border.Color,
        Type = border.Type
    };
}