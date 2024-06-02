using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertColumn(IXLColumn column) : ISpreadsheetAssertColumn
{
    public double Width => column.Width + 0.71062;

    public IEnumerable<ISpreadsheetAssertCell> Cells => column.CellsUsed().Select(x => new ClosedXmlAssertCell(x));
}
