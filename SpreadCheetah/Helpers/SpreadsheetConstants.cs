namespace SpreadCheetah.Helpers;

internal static class SpreadsheetConstants
{
    // These constants have been found by trial and error when working with embedded images.
    public const double DefaultColumnWidth = 8.11 + 0.77671875;
    public const double DefaultRowHeight = 18.375;

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
