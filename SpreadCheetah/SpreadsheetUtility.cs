using SpreadCheetah.Helpers;

namespace SpreadCheetah;

/// <summary>
/// Provides convenience methods related to working with spreadsheets.
/// </summary>
public static class SpreadsheetUtility
{
    /// <summary>
    /// Get the column name from a column number. E.g. column number 1 will return column name 'A'.
    /// </summary>
    public static string GetColumnName(int columnNumber)
    {
        if (columnNumber < 1 || columnNumber > SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber));

        if (columnNumber <= 26)
            return ((char)(columnNumber + 'A' - 1)).ToString();

        if (columnNumber <= 702)
        {
            var quotient = (columnNumber - 1) / 26;
            var remainder = (columnNumber - 1) % 26;
            char firstChar = (char)('A' - 1 + quotient);
            char secondChar = (char)('A' + remainder);
            return string.Format("{0}{1}", firstChar, secondChar);
        }
        else
        {
            var quotient1 = (columnNumber - 1) / 26;
            var remainder1 = (columnNumber - 1) % 26;
            var quotient2 = (quotient1 - 1) / 26;
            var remainder2 = (quotient1 - 1) % 26;
            char firstChar = (char)('A' - 1 + quotient2);
            char secondChar = (char)('A' + remainder2);
            char thirdChar = (char)('A' + remainder1);
            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
        }
    }
}
