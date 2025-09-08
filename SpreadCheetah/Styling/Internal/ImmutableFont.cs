using System.Drawing;

namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableFont(
    string Name,
    bool Bold,
    bool Italic,
    bool Strikethrough,
    Underline Underline,
    double Size,
    Color? Color)
{
    public static ImmutableFont From(Font font, DefaultFont? defaultFont) => new(
        Name: font.Name ?? defaultFont?.Name ?? DefaultFont.DefaultName,
        Bold: font.Bold,
        Italic: font.Italic,
        Strikethrough: font.Strikethrough,
        Underline: font.Underline,
        Size: font.ActualSize ?? defaultFont?.Size ?? DefaultFont.DefaultSize,
        Color: font.Color);

    public static ImmutableFont From(DefaultFont? defaultFont) => new(
        Name: defaultFont?.Name ?? DefaultFont.DefaultName,
        Bold: false,
        Italic: false,
        Strikethrough: false,
        Underline: Underline.None,
        Size: defaultFont?.Size ?? DefaultFont.DefaultSize,
        Color: null);
}
