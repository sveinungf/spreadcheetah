using SpreadCheetah.Helpers;

namespace SpreadCheetah;

/// <summary>
/// Provides convenience methods related to working with spreadsheets.
/// </summary>
public static class SpreadsheetUtility
{
    private static readonly string[] Alphabet = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

    /// <summary>
    /// Get the column name from a column number. E.g. column number 1 will return column name 'A'.
    /// </summary>
    public static string GetColumnName(int columnNumber)
    {
        var colNum = columnNumber - 1;
        if ((uint)colNum >= SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        if (colNum < 26)
            return Alphabet[colNum];

        var quotient1 = Math.DivRem(colNum, 26, out var remainder1);
        if (colNum < 702)
        {
#if NET8_0_OR_GREATER
            return new string(
            [
                (char)('A' - 1 + quotient1),
                (char)('A' + remainder1)
            ]);
#else
            return stackalloc char[2]
            {
                (char)('A' - 1 + quotient1),
                (char)('A' + remainder1)
            }.ToString();
#endif
        }

        var quotient2 = Math.DivRem(quotient1 - 1, 26, out var remainder2);
#if NET8_0_OR_GREATER
        return new string(
        [
            (char)('A' - 1 + quotient2),
            (char)('A' + remainder2),
            (char)('A' + remainder1)
        ]);
#else
        return stackalloc char[3]
        {
            (char)('A' - 1 + quotient2),
            (char)('A' + remainder2),
            (char)('A' + remainder1)
        }.ToString();
#endif
    }

    /// <summary>
    /// Try to write the UTF8 column name from a column number into the specified span. E.g. column number 1 results in column name 'A'.
    /// Returns <see langword="true"/> if the column name was written into the span, and <see langword="false"/> otherwise.
    /// </summary>
    public static bool TryGetColumnNameUtf8(int columnNumber, Span<byte> destination, out int bytesWritten)
    {
        var colNum = columnNumber - 1;
        if ((uint)colNum >= SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        if (colNum < 26)
        {
            if (destination.Length > 0)
            {
                destination[0] = (byte)(colNum + 'A');
                bytesWritten = 1;
                return true;
            }
        }
        else if (colNum < 702)
        {
            if (destination.Length > 1)
            {
                var quotient = Math.DivRem(colNum, 26, out var remainder);
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
                var quotient1 = Math.DivRem(colNum, 26, out var remainder1);
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
    /// Returns <see langword="true"/> if the column number was parsed successfully, and <see langword="false"/> otherwise.
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
