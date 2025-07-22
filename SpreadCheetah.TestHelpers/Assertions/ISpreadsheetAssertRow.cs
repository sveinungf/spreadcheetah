namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertRow
{
    IEnumerable<ISpreadsheetAssertCell> Cells { get; }
}
