namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyle
{
    ISpreadsheetAssertStyleBorder Border { get; }
    ISpreadsheetAssertStyleFill Fill { get; }
    ISpreadsheetAssertStyleFont Font { get; }
    ISpreadsheetAssertStyleNumberFormat NumberFormat { get; }
}
