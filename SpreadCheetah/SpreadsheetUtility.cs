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
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        if (columnNumber <= 26)
            return ((char)(columnNumber + 'A' - 1)).ToString();

        if (columnNumber <= 702)
        {
            Span<char> characters = stackalloc char[2];
            var quotient = Math.DivRem(columnNumber - 1, 26, out var remainder);
            characters[0] = (char)('A' - 1 + quotient);
            characters[1] = (char)('A' + remainder);
            return characters.ToString();
        }
        else
        {
            Span<char> characters = stackalloc char[3];
            var quotient1 = Math.DivRem(columnNumber - 1, 26, out var remainder1);
            var quotient2 = Math.DivRem(quotient1 - 1, 26, out var remainder2);
            characters[0] = (char)('A' - 1 + quotient2);
            characters[1] = (char)('A' + remainder2);
            characters[2] = (char)('A' + remainder1);
            return characters.ToString();
        }
    }
}
