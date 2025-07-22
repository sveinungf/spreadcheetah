namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertRow
{
    double Height { get; }
    ISpreadsheetAssertStyle Style { get; }
    IEnumerable<ISpreadsheetAssertCell> Cells { get; }
}
