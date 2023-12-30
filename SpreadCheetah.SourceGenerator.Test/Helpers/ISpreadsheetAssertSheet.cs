namespace SpreadCheetah.SourceGenerator.Test.Helpers;

internal interface ISpreadsheetAssertSheet : IDisposable
{
    ISpreadsheetAssertCell this[string columnName, int rowNumber] { get; }
    int CellCount { get; }
}
