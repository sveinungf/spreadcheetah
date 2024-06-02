namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertColumn
{
    double Width { get; }
    IEnumerable<ISpreadsheetAssertCell> Cells { get; }
}
