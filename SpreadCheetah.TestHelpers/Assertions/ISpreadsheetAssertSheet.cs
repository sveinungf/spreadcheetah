namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertSheet : IDisposable
{
    ISpreadsheetAssertCell this[string columnName, int rowNumber] { get; }
    int CellCount { get; }
    int RowCount { get; }

    IEnumerable<ISpreadsheetAssertCell> Column(string columnName);
    IEnumerable<ISpreadsheetAssertCell> Row(int rowNumber);
}
