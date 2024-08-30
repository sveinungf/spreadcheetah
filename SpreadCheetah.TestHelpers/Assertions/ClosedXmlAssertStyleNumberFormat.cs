using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal class ClosedXmlAssertStyleNumberFormat(IXLNumberFormat numberFormat)
    : ISpreadsheetAssertStyleNumberFormat
{
    public string? Format => numberFormat.Format is { Length: > 0 } format ? format : null;
}
