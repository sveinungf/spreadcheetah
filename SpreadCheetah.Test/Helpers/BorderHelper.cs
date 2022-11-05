using ClosedXML.Excel;
using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Helpers;

internal static class BorderHelper
{
    public static XLBorderStyleValues GetClosedXmlBorderStyle(this BorderStyle borderStyle)
    {
        return borderStyle switch
        {
            BorderStyle.None => XLBorderStyleValues.None,
            BorderStyle.Thin => XLBorderStyleValues.Thin,
            BorderStyle.Medium => XLBorderStyleValues.Medium,
            BorderStyle.Dashed => XLBorderStyleValues.Dashed,
            BorderStyle.Dotted => XLBorderStyleValues.Dotted,
            BorderStyle.Thick => XLBorderStyleValues.Thick,
            BorderStyle.DoubleLine => XLBorderStyleValues.Double,
            BorderStyle.Hair => XLBorderStyleValues.Hair,
            BorderStyle.MediumDashed => XLBorderStyleValues.MediumDashed,
            BorderStyle.DashDot => XLBorderStyleValues.DashDot,
            BorderStyle.MediumDashDot => XLBorderStyleValues.MediumDashDot,
            BorderStyle.DashDotDot => XLBorderStyleValues.DashDotDot,
            BorderStyle.MediumDashDotDot => XLBorderStyleValues.MediumDashDotDot,
            BorderStyle.SlantDashDot => XLBorderStyleValues.SlantDashDot,
            _ => throw new NotImplementedException(),
        };
    }
}
