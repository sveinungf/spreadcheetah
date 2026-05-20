using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlWorksheetColumn(IXLColumn column) : IWorksheetColumn
{
    public bool Hidden => column.IsHidden;

    public double Width => column.Width + 0.71062;

    public IStyle Style => ClosedXmlStyle.Create(column.Style);

    public IEnumerable<IWorksheetCell> Cells => column.CellsUsed().Select(x => new ClosedXmlWorksheetCell(x));
}
