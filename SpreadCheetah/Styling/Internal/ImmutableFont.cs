using System.Drawing;

namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableFont(
    string? Name,
    bool Bold,
    bool Italic,
    bool Strikethrough,
    double Size,
    Color? Color)
{
    public static ImmutableFont From(Font font) => new(
        Name: font.Name,
        Bold: font.Bold,
        Italic: font.Italic,
        Strikethrough: font.Strikethrough,
        Size: font.Size,
        Color: font.Color);
}
