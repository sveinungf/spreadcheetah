using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertStyle(IXLStyle style) : ISpreadsheetAssertStyle
{
    public ISpreadsheetAssertStyleFont Font => new ClosedXmlAssertStyleFont(style.Font);

    public ISpreadsheetAssertStyleNumberFormat NumberFormat => new ClosedXmlAssertStyleNumberFormat(style.NumberFormat);
}
