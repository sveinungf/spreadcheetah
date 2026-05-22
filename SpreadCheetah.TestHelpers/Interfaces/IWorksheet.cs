namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IWorksheet : IDisposable
{
    string Name { get; }

    IWorksheetCell this[string reference] { get; }
    IWorksheetCell this[string columnName, int rowNumber] { get; }
    int CellCount { get; }
    int RowCount { get; }
    int? MaxRowOutlineLevel { get; }

    IEnumerable<IWorksheetCell> AllCells();
    IWorksheetColumn Column(string columnName);
    IReadOnlyList<IWorksheetColumn> Columns { get; }
    IWorksheetRow Row(int rowNumber);

    IReadOnlyList<IConditionalFormatRule> ConditionalFormatRules { get; }
    IReadOnlyList<IWorksheetTable> Tables { get; }
}
