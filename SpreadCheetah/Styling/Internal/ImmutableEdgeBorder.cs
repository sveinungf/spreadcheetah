using SpreadCheetah.ConditionalFormatting;
using System.Drawing;
using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableEdgeBorder
{
    public required BorderStyle BorderStyle { get; init; }
    public required Color? Color { get; init; }

    public static ImmutableEdgeBorder From(EdgeBorder border) => new()
    {
        BorderStyle = border.BorderStyle,
        Color = border.Color
    };

    public static ImmutableEdgeBorder? From(ConditionalFormatEdgeBorder border)
    {
        if (border is { BorderStyle: ConditionalFormatBorderStyle.None, Color: null })
            return null;

        return new()
        {
            BorderStyle = (BorderStyle)border.BorderStyle,
            Color = border.Color
        };
    }
}
