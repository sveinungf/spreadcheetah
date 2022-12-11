using System.Runtime.InteropServices;

namespace SpreadCheetah.Styling.Internal;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct ImmutableAlignment(
    HorizontalAlignment Horizontal,
    VerticalAlignment Vertical,
    bool WrapText,
    int Indent)
{
    public static ImmutableAlignment From(Alignment alignment) => new(
        Horizontal: alignment.Horizontal,
        Vertical: alignment.Vertical,
        WrapText: alignment.WrapText,
        Indent: alignment.Indent);
}
