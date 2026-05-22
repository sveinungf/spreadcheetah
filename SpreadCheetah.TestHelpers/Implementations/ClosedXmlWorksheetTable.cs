using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlWorksheetTable(IXLTable table) : IWorksheetTable
{
    public string CellRangeReference => table.RangeAddress.ToString() ?? "";
    public string Name => table.Name;
    public string TableStyle => table.Theme.Name;
    public bool BandedColumns => table.ShowColumnStripes;
    public bool BandedRows => table.ShowRowStripes;
    public bool ShowAutoFilter => table.ShowAutoFilter;
    public bool ShowHeaderRow => table.ShowHeaderRow;
    public bool ShowTotalRow => table.ShowTotalsRow;

    public IReadOnlyList<IWorksheetTableColumn> Columns => field ??=
        [.. table.Fields.Select(x => new ClosedXmlWorksheetTableColumn(x))];
}
