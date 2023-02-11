namespace SpreadCheetah.Helpers;

internal static class SpreadsheetConstants
{
    // We are not expecting more than 9,999 sheets in one file
    public const int SheetCountMaxDigits = 4;

    // Limitation set by the Excel specification
    public const int MaxNumberOfColumns = 16384;

    // Limitation set by the Excel specification
    public const int MaxNumberOfRows = 1048576;

    // Limitation set by the Excel specification
    public const int MaxNumberOfDataValidations = 65534;
}
