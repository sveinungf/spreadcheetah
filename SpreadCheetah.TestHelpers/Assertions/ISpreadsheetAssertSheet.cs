namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertSheet : IDisposable
{
    ISpreadsheetAssertCell this[string columnName, int rowNumber] { get; }
    IEnumerable<ISpreadsheetAssertCell> this[string columnName] { get; }
    int CellCount { get; }
    int RowCount { get; }

    IEnumerable<ISpreadsheetAssertCell> Row(int rowNumber);
}
