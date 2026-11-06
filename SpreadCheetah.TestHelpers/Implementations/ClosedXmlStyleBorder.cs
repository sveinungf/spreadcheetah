using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Interfaces;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed record ClosedXmlStyleBorder : IStyleBorder
{
    public required Color BottomColor { get; init; }
    public required BorderStyle BottomStyle { get; init; }
    public required Color DiagonalColor { get; init; }
    public required BorderStyle DiagonalStyle { get; init; }
    public required Color LeftColor { get; init; }
    public required BorderStyle LeftStyle { get; init; }
    public required Color RightColor { get; init; }
    public required BorderStyle RightStyle { get; init; }
    public required Color TopColor { get; init; }
    public required BorderStyle TopStyle { get; init; }

    public static ClosedXmlStyleBorder Create(IXLBorder border)
    {
        return new()
        {
            BottomColor = border.BottomBorderColor.Color,
            BottomStyle = Map(border.BottomBorder),
            DiagonalColor = border.DiagonalBorderColor.Color,
            DiagonalStyle = Map(border.DiagonalBorder),
            LeftColor = border.LeftBorderColor.Color,
            LeftStyle = Map(border.LeftBorder),
            RightColor = border.RightBorderColor.Color,
            RightStyle = Map(border.RightBorder),
            TopColor = border.TopBorderColor.Color,
            TopStyle = Map(border.TopBorder)
        };
    }

    public static ClosedXmlStyleBorder Create(Border border)
    {
        return new()
        {
            BottomColor = MapColor(border.Bottom),
            BottomStyle = border.Bottom.BorderStyle,
            DiagonalColor = MapColor(border.Diagonal),
            DiagonalStyle = border.Diagonal.BorderStyle,
            LeftColor = MapColor(border.Left),
            LeftStyle = border.Left.BorderStyle,
            RightColor = MapColor(border.Right),
            RightStyle = border.Right.BorderStyle,
            TopColor = MapColor(border.Top),
            TopStyle = border.Top.BorderStyle
        };
    }

    private static Color MapColor(EdgeBorder border) => MapColor(border.BorderStyle, border.Color);
    private static Color MapColor(DiagonalBorder border) => MapColor(border.BorderStyle, border.Color);

    private static Color MapColor(BorderStyle borderStyle, Color? color)
    {
        return (borderStyle, color) switch
        {
            (BorderStyle.None, _) => Color.Black,
            (_, { } actualColor) => actualColor,
            _ => Color.Black
        };
    }

    private static BorderStyle Map(XLBorderStyleValues value)
    {
        return value switch
        {
            XLBorderStyleValues.None => BorderStyle.None,
            XLBorderStyleValues.Thin => BorderStyle.Thin,
            XLBorderStyleValues.Medium => BorderStyle.Medium,
            XLBorderStyleValues.Dashed => BorderStyle.Dashed,
            XLBorderStyleValues.Dotted => BorderStyle.Dotted,
            XLBorderStyleValues.Thick => BorderStyle.Thick,
            XLBorderStyleValues.Double => BorderStyle.DoubleLine,
            XLBorderStyleValues.Hair => BorderStyle.Hair,
            XLBorderStyleValues.MediumDashed => BorderStyle.MediumDashed,
            XLBorderStyleValues.DashDot => BorderStyle.DashDot,
            XLBorderStyleValues.MediumDashDot => BorderStyle.MediumDashDot,
            XLBorderStyleValues.DashDotDot => BorderStyle.DashDotDot,
            XLBorderStyleValues.MediumDashDotDot => BorderStyle.MediumDashDotDot,
            XLBorderStyleValues.SlantDashDot => BorderStyle.SlantDashDot,
            _ => throw new ArgumentOutOfRangeException(nameof(value), value, null)
        };
    }
}
