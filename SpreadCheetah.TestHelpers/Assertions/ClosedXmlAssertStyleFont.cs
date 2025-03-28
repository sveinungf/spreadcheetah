using ClosedXML.Excel;
using SpreadCheetah.Styling;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertStyleFont(IXLFont font) : ISpreadsheetAssertStyleFont
{
    public bool Bold => font.Bold;
    public bool Italic => font.Italic;
    public bool Strikethrough => font.Strikethrough;

    public Underline Underline => font.Underline switch
    {
        XLFontUnderlineValues.None => Underline.None,
        XLFontUnderlineValues.Single => Underline.Single,
        XLFontUnderlineValues.SingleAccounting => Underline.SingleAccounting,
        XLFontUnderlineValues.Double => Underline.Double,
        XLFontUnderlineValues.DoubleAccounting => Underline.DoubleAccounting,
        _ => throw new ArgumentOutOfRangeException(nameof(font), font.Underline, null),
    };
}
