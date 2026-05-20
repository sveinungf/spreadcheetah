using ClosedXML.Excel;
using SpreadCheetah.Styling;
using System.Drawing;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertStyleBorder(IXLBorder border)
    : ISpreadsheetAssertStyleBorder
{
    public Color BottomColor => border.BottomBorderColor.Color;
    public BorderStyle BottomStyle => Map(border.BottomBorder);
    public Color DiagonalColor => border.DiagonalBorderColor.Color;
    public BorderStyle DiagonalStyle => Map(border.DiagonalBorder);
    public Color LeftColor => border.LeftBorderColor.Color;
    public BorderStyle LeftStyle => Map(border.LeftBorder);
    public Color RightColor => border.RightBorderColor.Color;
    public BorderStyle RightStyle => Map(border.RightBorder);
    public Color TopColor => border.TopBorderColor.Color;
    public BorderStyle TopStyle => Map(border.TopBorder);

    private static BorderStyle Map(XLBorderStyleValues value) => value switch
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
