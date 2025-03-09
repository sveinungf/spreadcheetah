using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertTable(IXLTable table) : ISpreadsheetAssertTable
{
    public string CellRangeReference => table.RangeAddress.ToString();
    public string Name => table.Name;
    public string TableStyle => table.Theme.Name;
    public bool BandedColumns => table.ShowColumnStripes;
    public bool BandedRows => table.ShowRowStripes;
    public bool ShowAutoFilter => table.ShowAutoFilter;
    public bool ShowHeaderRow => table.ShowHeaderRow;
    public bool ShowTotalRow => table.ShowTotalsRow;

    private IReadOnlyList<ISpreadsheetAssertTableColumn>? _columns;
    public IReadOnlyList<ISpreadsheetAssertTableColumn> Columns => _columns ??=
        [.. table.Fields.Select(x => new ClosedXmlAssertTableColumn(x))];
}
