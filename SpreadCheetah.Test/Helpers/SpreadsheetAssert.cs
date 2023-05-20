using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SpreadCheetah.Styling;
using System.Drawing;
using Xunit;

namespace SpreadCheetah.Test.Helpers;

internal static class SpreadsheetAssert
{
    public static void Valid(Stream stream)
    {
        stream.Position = 0;

        using var document = SpreadsheetDocument.Open(stream, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(document);
        Assert.Empty(errors);

        stream.Position = 0;
    }

    public static void EquivalentStyle(Style style, IXLStyle closedXmlStyle)
    {
        Assert.Equal(style.Alignment.Horizontal.GetClosedXmlHorizontalAlignment(), closedXmlStyle.Alignment.Horizontal);
        Assert.Equal(style.Alignment.Indent, closedXmlStyle.Alignment.Indent);
        Assert.Equal(style.Alignment.Vertical.GetClosedXmlVerticalAlignment(), closedXmlStyle.Alignment.Vertical);
        Assert.Equal(style.Alignment.WrapText, closedXmlStyle.Alignment.WrapText);

        AssertEdgeBorder(style.Border.Bottom, closedXmlStyle.Border.BottomBorder, closedXmlStyle.Border.BottomBorderColor);
        AssertEdgeBorder(style.Border.Left, closedXmlStyle.Border.LeftBorder, closedXmlStyle.Border.LeftBorderColor);
        AssertEdgeBorder(style.Border.Right, closedXmlStyle.Border.RightBorder, closedXmlStyle.Border.RightBorderColor);
        AssertEdgeBorder(style.Border.Top, closedXmlStyle.Border.TopBorder, closedXmlStyle.Border.TopBorderColor);
        AssertDiagonalBorder(style.Border.Diagonal, closedXmlStyle.Border);

        Assert.Equal(style.Fill.Color?.ToArgb() ?? 16777215, closedXmlStyle.Fill.BackgroundColor.Color.ToArgb());

        Assert.Equal(style.Font.Bold, closedXmlStyle.Font.Bold);
        Assert.Equal(style.Font.Italic, closedXmlStyle.Font.Italic);
        Assert.Equal(style.Font.Name ?? "Calibri", closedXmlStyle.Font.FontName);
        Assert.Equal(style.Font.Size, closedXmlStyle.Font.FontSize);
        Assert.Equal(style.Font.Strikethrough, closedXmlStyle.Font.Strikethrough);
        AssertColor(style.Font.Color, closedXmlStyle.Font.FontColor);

        Assert.Equal(style.NumberFormat ?? "", closedXmlStyle.NumberFormat.Format);
    }


    private static void AssertEdgeBorder(EdgeBorder border, XLBorderStyleValues style, XLColor color)
    {
        Assert.Equal(border.BorderStyle.GetClosedXmlBorderStyle(), style);

        if (border.BorderStyle != BorderStyle.None)
            AssertColor(border.Color, color);
    }

    private static void AssertDiagonalBorder(DiagonalBorder border, IXLBorder closedXmlBorder)
    {
        Assert.Equal(border.BorderStyle.GetClosedXmlBorderStyle(), closedXmlBorder.DiagonalBorder);

        if (border.BorderStyle != BorderStyle.None)
        {
            AssertColor(border.Color, closedXmlBorder.DiagonalBorderColor);

            Assert.Equal(border.Type.HasFlag(DiagonalBorderType.DiagonalUp), closedXmlBorder.DiagonalUp);
            Assert.Equal(border.Type.HasFlag(DiagonalBorderType.DiagonalDown), closedXmlBorder.DiagonalDown);
        }
    }

    private static void AssertColor(Color? color, XLColor closedXmlColor)
    {
        color ??= Color.FromArgb(255, 0, 0, 0);
        Assert.Equal(XLColor.FromColor(color.Value), closedXmlColor);
    }
}
