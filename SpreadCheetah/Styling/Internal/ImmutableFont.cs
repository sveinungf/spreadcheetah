using SpreadCheetah.ConditionalFormatting;
using System.Drawing;

namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableFont
{
    public required string? Name { get; init; }
    public required bool Bold { get; init; }
    public required bool Italic { get; init; }
    public required bool Strikethrough { get; init; }
    public required Underline Underline { get; init; }
    public required double? Size { get; init; }
    public required Color? Color { get; init; }

    public static ImmutableFont From(Font font, DefaultFont? defaultFont) => new()
    {
        Bold = font.Bold,
        Color = font.Color,
        Italic = font.Italic,
        Name = font.Name ?? defaultFont?.Name ?? DefaultFont.DefaultName,
        Size = font.ActualSize ?? defaultFont?.Size ?? DefaultFont.DefaultSize,
        Strikethrough = font.Strikethrough,
        Underline = font.Underline
    };

    public static ImmutableFont From(ConditionalFormatFont? font) => new()
    {
        Bold = font?.Bold ?? false,
        Color = font?.Color,
        Italic = font?.Italic ?? false,
        Name = null,
        Size = null,
        Strikethrough = font?.Strikethrough ?? false,
        Underline = (Underline)(font?.Underline ?? ConditionalFormatUnderline.None)
    };

    public static ImmutableFont From(DefaultFont? defaultFont) => new()
    {
        Bold = false,
        Color = null,
        Italic = false,
        Name = defaultFont?.Name ?? DefaultFont.DefaultName,
        Size = defaultFont?.Size ?? DefaultFont.DefaultSize,
        Strikethrough = false,
        Underline = Underline.None
    };
}
