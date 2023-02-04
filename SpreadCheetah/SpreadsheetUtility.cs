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
        if (columnNumber < 1 || columnNumber > SpreadsheetConstants.MaxNumberOfColumns) //Excel columns are one-based (one = 'A')
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber));

        if (columnNumber <= 26) //one character
            return ((char)(columnNumber + 'A' - 1)).ToString();

        if (columnNumber <= 702) //two characters
        {
            char firstChar = (char)((columnNumber - 1) / 26 + 'A' - 1);
            char secondChar = (char)(columnNumber % 26 + 'A' - 1);

            if (secondChar == '@') //Excel is one-based, but modulo operations are zero-based
                secondChar = 'Z'; //convert one-based to zero-based

            return string.Format("{0}{1}", firstChar, secondChar);
        }
        else //three characters
        {
            char firstChar = (char)((columnNumber - 1) / 702 + 'A' - 1);
            char secondChar = (char)((columnNumber - 1) / 26 % 26 + 'A' - 1);
            char thirdChar = (char)(columnNumber % 26 + 'A' - 1);

            if (thirdChar == '@') //Excel is one-based, but modulo operations are zero-based
                thirdChar = 'Z'; //convert one-based to zero-based

            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
        }
    }
}
