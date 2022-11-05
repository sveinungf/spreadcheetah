using System.Drawing;
using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableDiagonalBorder(
    BorderStyle BorderStyle,
    Color? Color,
    DiagonalBorderType Type)
{
    public static ImmutableDiagonalBorder From(DiagonalBorder border) => new(
        BorderStyle: border.BorderStyle,
        Color: border.Color,
        Type: border.Type);
}