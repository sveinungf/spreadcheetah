namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertRow
{
    double Height { get; }
    int OutlineLevel { get; }
    bool Hidden { get; }
    bool Collapsed { get; }

    ISpreadsheetAssertStyle Style { get; }
    IEnumerable<ISpreadsheetAssertCell> Cells { get; }
}
