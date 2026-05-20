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
