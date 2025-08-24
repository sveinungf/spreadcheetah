namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertColumn
{
    double Width { get; }
    ISpreadsheetAssertStyle Style { get; }
    IEnumerable<ISpreadsheetAssertCell> Cells { get; }
}
