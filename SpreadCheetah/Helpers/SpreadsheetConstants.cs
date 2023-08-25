namespace SpreadCheetah.Helpers;

internal static class SpreadsheetConstants
{
    public const int MaxNoteTextLength = 32768;

    // Limitation set by the Excel specification
    public const int MaxNumberOfColumns = 16384;

    // Limitation set by the Excel specification
    public const int MaxNumberOfRows = 1048576;

    // Limitation set by the Excel specification
    public const int MaxNumberOfDataValidations = 65534;
}
