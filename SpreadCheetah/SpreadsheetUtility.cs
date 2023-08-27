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

    /// <summary>
    /// Try to write the UTF8 column name from a column number into the specified span. E.g. column number 1 results in column name 'A'.
    /// Returns <c>true</c> if the column name was written into the span, and <c>false</c> otherwise.
    /// </summary>
    public static bool TryGetColumnNameUtf8(int columnNumber, Span<byte> destination, out int bytesWritten)
    {
        if (columnNumber < 1 || columnNumber > SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        if (columnNumber <= 26)
        {
            if (destination.Length > 0)
            {
                destination[0] = (byte)(columnNumber + 'A' - 1);
                bytesWritten = 1;
                return true;
            }
        }
        else if (columnNumber <= 702)
        {
            if (destination.Length > 1)
            {
                var quotient = Math.DivRem(columnNumber - 1, 26, out var remainder);
                destination[0] = (byte)('A' - 1 + quotient);
                destination[1] = (byte)('A' + remainder);
                bytesWritten = 2;
                return true;
            }
        }
        else
        {
            if (destination.Length > 2)
            {
                var quotient1 = Math.DivRem(columnNumber - 1, 26, out var remainder1);
                var quotient2 = Math.DivRem(quotient1 - 1, 26, out var remainder2);
                destination[0] = (byte)('A' - 1 + quotient2);
                destination[1] = (byte)('A' + remainder2);
                destination[2] = (byte)('A' + remainder1);
                bytesWritten = 3;
                return true;
            }
        }

        bytesWritten = 0;
        return false;
    }

    /// <summary>
    /// Try to parse the column number from a column name. E.g. column name 'A' can be parsed to column number 1.
    /// Returns <c>true</c> if the column number was parsed successfully, and <c>false</c> otherwise.
    /// </summary>
    public static bool TryParseColumnName(ReadOnlySpan<char> columnName, out int columnNumber)
    {
        if (columnName is [var a] && IsValid(a))
        {
            columnNumber = ToNumber(a);
            return true;
        }

        if (columnName is [var b0, var b1] && IsValid(b0) && IsValid(b1))
        {
            columnNumber = ToNumber(b0) * 26 + ToNumber(b1);
            return true;
        }

        if (columnName is [var c0, var c1, var c2] && IsValid(c0) && IsValid(c1) && IsValid(c2))
        {
            columnNumber = (ToNumber(c0) * 26 + ToNumber(c1)) * 26 + ToNumber(c2);
            if (columnNumber <= SpreadsheetConstants.MaxNumberOfColumns)
                return true;
        }

        columnNumber = 0;
        return false;

        static bool IsValid(char c) => (uint)(c - 'A') <= ('Z' - 'A');
        static int ToNumber(char c) => c - 'A' + 1;
    }
}
