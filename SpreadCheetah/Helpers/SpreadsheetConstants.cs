namespace SpreadCheetah.Helpers;

internal static class SpreadsheetConstants
{
    public const double DefaultColumnWidth = 8.11;
    public const double DefaultRowHeight = 14.4;

    // Limitation set by SpreadCheetah
    public const int MaxImageDimension = ushort.MaxValue;

    // Limitation set by SpreadCheetah
    public const int MaxNoteTextLength = 32768;

    // Limitation set by the Excel specification
    public const int MaxNumberOfColumns = 16384;

    // Limitation set by the Excel specification
    public const int MaxNumberOfRows = 1048576;

    // Limitation set by the Excel specification
    public const int MaxNumberOfDataValidations = 65534;
}
