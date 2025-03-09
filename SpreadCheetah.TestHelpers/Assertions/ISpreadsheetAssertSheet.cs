namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertSheet : IDisposable
{
    string Name { get; }

    ISpreadsheetAssertCell this[string reference] { get; }
    ISpreadsheetAssertCell this[string columnName, int rowNumber] { get; }
    int CellCount { get; }
    int RowCount { get; }

    ISpreadsheetAssertColumn Column(string columnName);
    IReadOnlyList<ISpreadsheetAssertColumn> Columns { get; }
    IEnumerable<ISpreadsheetAssertCell> Row(int rowNumber);

    IReadOnlyList<ISpreadsheetAssertTable> Tables { get; }
}
