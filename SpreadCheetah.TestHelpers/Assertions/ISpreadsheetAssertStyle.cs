namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyle
{
    ISpreadsheetAssertStyleFont Font { get; }
    ISpreadsheetAssertStyleNumberFormat NumberFormat { get; }
}
