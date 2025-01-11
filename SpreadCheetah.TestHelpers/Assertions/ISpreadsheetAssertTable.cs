namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertTable
{
    string CellRangeReference { get; }
    string Name { get; }
    string TableStyle { get; }
    bool BandedColumns { get; }
    bool BandedRows { get; }
    bool ShowAutoFilter { get; }
    bool ShowHeaderRow { get; }
    bool ShowTotalRow { get; }
    IReadOnlyList<ISpreadsheetAssertTableColumn> Columns { get; }
}
