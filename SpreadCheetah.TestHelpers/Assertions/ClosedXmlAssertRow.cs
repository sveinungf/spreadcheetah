using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertRow(IXLRow row) : ISpreadsheetAssertRow
{
    public double Height => row.Height;

    public int OutlineLevel => row.OutlineLevel;

    public ISpreadsheetAssertStyle Style => new ClosedXmlAssertStyle(row.Style);

    public bool Hidden => row.IsHidden;

    public bool Collapsed => throw new NotImplementedException("Currently, the closedxml property Collapsed is not in the interface, but it is in the implementation. If the behavior changes, the code must be updated.");

    public IEnumerable<ISpreadsheetAssertCell> Cells => row.IsEmpty()
        ? []
        : row.Cells(false).Select(x => new ClosedXmlAssertCell(x));
}
