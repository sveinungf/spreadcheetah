namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IWorksheetTable
{
    string CellRangeReference { get; }
    string Name { get; }
    string TableStyle { get; }
    bool BandedColumns { get; }
    bool BandedRows { get; }
    bool ShowAutoFilter { get; }
    bool ShowHeaderRow { get; }
    bool ShowTotalRow { get; }
    IReadOnlyList<IWorksheetTableColumn> Columns { get; }
}
