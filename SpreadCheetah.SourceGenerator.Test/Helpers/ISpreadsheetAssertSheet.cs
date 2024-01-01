namespace SpreadCheetah.SourceGenerator.Test.Helpers;

internal interface ISpreadsheetAssertSheet : IDisposable
{
    ISpreadsheetAssertCell this[string columnName, int rowNumber] { get; }
    IEnumerable<ISpreadsheetAssertCell> this[string columnName] { get; }
    int CellCount { get; }
    int RowCount { get; }
}
