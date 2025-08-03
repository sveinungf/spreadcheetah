namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyle
{
    ISpreadsheetAssertStyleFill Fill { get; }
    ISpreadsheetAssertStyleFont Font { get; }
    ISpreadsheetAssertStyleNumberFormat NumberFormat { get; }
}
