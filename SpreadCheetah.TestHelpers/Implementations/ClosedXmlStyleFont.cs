using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Interfaces;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed record ClosedXmlStyleFont : IStyleFont
{
    public required bool Bold { get; init; }
    public required bool Italic { get; init; }
    public required bool Strikethrough { get; init; }
    public required double Size { get; init; }
    public required string Name { get; init; }
    public required Color Color { get; init; }
    public required Underline Underline { get; init; }

    public static ClosedXmlStyleFont Create(IXLFont font)
    {
        return new()
        {
            Bold = font.Bold,
            Italic = font.Italic,
            Strikethrough = font.Strikethrough,
            Size = font.FontSize,
            Name = font.FontName,
            Color = font.FontColor.Color,
            Underline = Map(font.Underline)
        };
    }

    public static ClosedXmlStyleFont Create(Font font)
    {
        return new()
        {
            Bold = font.Bold,
            Italic = font.Italic,
            Strikethrough = font.Strikethrough,
            Size = font.Size,
            Name = font.Name ?? "Calibri",
            Color = font.Color ?? Color.Black,
            Underline = font.Underline
        };
    }

    private static Underline Map(XLFontUnderlineValues value)
    {
        return value switch
        {
            XLFontUnderlineValues.None => Underline.None,
            XLFontUnderlineValues.Single => Underline.Single,
            XLFontUnderlineValues.SingleAccounting => Underline.SingleAccounting,
            XLFontUnderlineValues.Double => Underline.Double,
            XLFontUnderlineValues.DoubleAccounting => Underline.DoubleAccounting,
            _ => throw new ArgumentOutOfRangeException(nameof(value), value, null),
        };
    }
}
