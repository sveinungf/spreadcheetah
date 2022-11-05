using System.Drawing;
using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableEdgeBorder(
    BorderStyle BorderStyle,
    Color? Color)
{
    public static ImmutableEdgeBorder From(EdgeBorder border) => new(
        BorderStyle: border.BorderStyle,
        Color: border.Color);
}
