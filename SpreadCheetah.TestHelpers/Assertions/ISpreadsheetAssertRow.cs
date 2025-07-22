namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertRow
{
    double Height { get; }
    IEnumerable<ISpreadsheetAssertCell> Cells { get; }
}
