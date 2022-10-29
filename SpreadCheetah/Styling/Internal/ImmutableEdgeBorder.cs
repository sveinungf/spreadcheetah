using System.Drawing;

namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableEdgeBorder(
    BorderStyle BorderStyle,
    Color? Color)
{
    public static ImmutableEdgeBorder From(EdgeBorder border) => new(
        BorderStyle: border.BorderStyle,
        Color: border.Color);
}
