namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertColumn
{
    bool Hidden { get; }
    double Width { get; }
    ISpreadsheetAssertStyle Style { get; }
    IEnumerable<ISpreadsheetAssertCell> Cells { get; }
}
