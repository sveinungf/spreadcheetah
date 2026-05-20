namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertStyle
{
    ISpreadsheetAssertStyleAlignment Alignment { get; }
    ISpreadsheetAssertStyleBorder Border { get; }
    ISpreadsheetAssertStyleFill Fill { get; }
    ISpreadsheetAssertStyleFont Font { get; }
    ISpreadsheetAssertStyleNumberFormat NumberFormat { get; }
}
