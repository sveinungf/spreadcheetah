namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertRow
{
    double Height { get; }
    int OutlineLevel { get; }
    ISpreadsheetAssertStyle Style { get; }
    IEnumerable<ISpreadsheetAssertCell> Cells { get; }
}
