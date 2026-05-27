using System.Drawing;
using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableEdgeBorder
{
    public required BorderStyle BorderStyle { get; init; }
    public required Color? Color { get; init; }

    public bool IsDefault => this is
    {
        BorderStyle: BorderStyle.None,
        Color: null
    };

    public static ImmutableEdgeBorder From(EdgeBorder border) => new()
    {
        BorderStyle = border.BorderStyle,
        Color = border.Color
    };
}
