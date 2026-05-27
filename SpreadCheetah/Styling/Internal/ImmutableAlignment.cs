using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableAlignment
{
    public required HorizontalAlignment Horizontal { get; init; }
    public required VerticalAlignment Vertical { get; init; }
    public required bool WrapText { get; init; }
    public required int Indent { get; init; }

    public bool IsDefault => this is
    {
        Horizontal: HorizontalAlignment.None,
        Vertical: VerticalAlignment.Bottom,
        WrapText: false,
        Indent: 0
    };

    public static ImmutableAlignment From(Alignment alignment) => new()
    {
        Horizontal = alignment.Horizontal,
        Vertical = alignment.Vertical,
        WrapText = alignment.WrapText,
        Indent = alignment.Indent
    };
}
