using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertStyleFont(IXLFont font) : ISpreadsheetAssertStyleFont
{
    public bool Bold => font.Bold;
    public bool Italic => font.Italic;
}
