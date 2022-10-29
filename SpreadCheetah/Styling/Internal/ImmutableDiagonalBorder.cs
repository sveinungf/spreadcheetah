using System.Drawing;

namespace SpreadCheetah.Styling.Internal;

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