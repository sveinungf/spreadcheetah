using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlWorksheetRow(IXLRow row) : IWorksheetRow
{
    public double Height => row.Height;

    public int OutlineLevel => row.OutlineLevel;

    public IStyle Style => ClosedXmlStyle.Create(row.Style);

    public IEnumerable<IWorksheetCell> Cells => row.IsEmpty()
        ? []
        : row.Cells(false).Select(x => new ClosedXmlWorksheetCell(x));
}
