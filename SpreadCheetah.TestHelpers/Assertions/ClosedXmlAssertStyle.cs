using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertStyle(IXLStyle style) : ISpreadsheetAssertStyle
{
    public ISpreadsheetAssertStyleAlignment Alignment => ClosedXmlAssertStyleAlignment.Create(style.Alignment);

    public ISpreadsheetAssertStyleBorder Border => new ClosedXmlAssertStyleBorder(style.Border);

    public ISpreadsheetAssertStyleFill Fill => new ClosedXmlAssertStyleFill(style.Fill);

    public ISpreadsheetAssertStyleFont Font => new ClosedXmlAssertStyleFont(style.Font);

    public ISpreadsheetAssertStyleNumberFormat NumberFormat => new ClosedXmlAssertStyleNumberFormat(style.NumberFormat);
}
