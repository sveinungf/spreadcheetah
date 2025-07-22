using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertRow(IXLRow row) : ISpreadsheetAssertRow
{
    public IEnumerable<ISpreadsheetAssertCell> Cells => row.IsEmpty()
        ? []
        : row.Cells(false).Select(x => new ClosedXmlAssertCell(x));
}
