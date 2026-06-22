using SpreadCheetah.Helpers;
using System.Globalization;
using System.Text;

namespace SpreadCheetah.Formulas;

internal static class R1C1FormulaConverter
{
    private enum ReferenceKind
    {
        Cell,
        WholeRow,
        WholeColumn
    }

    /// <summary>
    /// Converts a formula in the R1C1 reference style to the A1 reference style.
    /// <paramref name="row"/> and <paramref name="column"/> are the 1-based position of the cell that the formula belongs to,
    /// which is used as the anchor for relative references.
    /// </summary>
    public static string ToA1(string formula, int row, int column)
    {
        var sb = new StringBuilder(formula.Length);
        var i = 0;
        var n = formula.Length;

        while (i < n)
        {
            var c = formula[i];

            if (c is '"' or '\'')
            {
                i = AppendDelimited(formula, i, c, sb);
                continue;
            }

            var leftBoundary = i == 0 || !IsWordChar(formula[i - 1]);
            if (leftBoundary && c is 'R' or 'r' or 'C' or 'c'
                && TryParseReference(formula, i, row, column, out var length, out var piece, out var kind))
            {
                var afterFirst = i + length;
                var rightChar = afterFirst < n ? formula[afterFirst] : '\0';

                // Combine into a range when followed by ':' and another reference.
                if (rightChar == ':'
                    && TryParseReference(formula, afterFirst + 1, row, column, out var length2, out var piece2, out _))
                {
                    var afterSecond = afterFirst + 1 + length2;
                    var rightChar2 = afterSecond < n ? formula[afterSecond] : '\0';
                    if (!IsWordChar(rightChar2) && rightChar2 is not ('[' or '!'))
                    {
                        sb.Append(piece).Append(':').Append(piece2);
                        i = afterSecond;
                        continue;
                    }
                }

                if (!IsWordChar(rightChar) && rightChar is not ('[' or '!'))
                {
                    AppendStandalone(sb, piece, kind);
                    i = afterFirst;
                    continue;
                }
            }

            sb.Append(c);
            i++;
        }

        return sb.ToString();
    }

    private static void AppendStandalone(StringBuilder sb, string piece, ReferenceKind kind)
    {
        // A whole row or column reference in the A1 notation always has the form "X:X".
        if (kind is ReferenceKind.Cell)
            sb.Append(piece);
        else
            sb.Append(piece).Append(':').Append(piece);
    }

    // Appends a string literal ("...") or a quoted sheet name ('...'), including the surrounding delimiters.
    // The delimiter can be escaped by doubling it.
    private static int AppendDelimited(string s, int i, char delimiter, StringBuilder sb)
    {
        var n = s.Length;
        sb.Append(s[i]);
        i++;

        while (i < n)
        {
            var c = s[i];
            sb.Append(c);
            i++;

            if (c == delimiter)
            {
                if (i < n && s[i] == delimiter)
                {
                    sb.Append(delimiter);
                    i++;
                    continue;
                }

                break;
            }
        }

        return i;
    }

    private static bool TryParseReference(string s, int start, int anchorRow, int anchorColumn, out int length, out string piece, out ReferenceKind kind)
    {
        length = 0;
        piece = "";
        kind = ReferenceKind.Cell;

        var n = s.Length;
        if (start >= n)
            return false;

        var j = start;
        var c = s[j];

        var hasRow = false;
        var hasColumn = false;
        bool rowRelative = true, columnRelative = true;
        int rowValue = 0, columnValue = 0;

        if (c is 'R' or 'r')
        {
            j++;
            hasRow = true;
            if (!TryParseAxis(s, ref j, out rowRelative, out rowValue))
                return false;

            if (j < n && s[j] is 'C' or 'c')
            {
                j++;
                hasColumn = true;
                if (!TryParseAxis(s, ref j, out columnRelative, out columnValue))
                    return false;
            }
        }
        else if (c is 'C' or 'c')
        {
            j++;
            hasColumn = true;
            if (!TryParseAxis(s, ref j, out columnRelative, out columnValue))
                return false;
        }
        else
        {
            return false;
        }

        length = j - start;

        if (hasRow && hasColumn)
        {
            kind = ReferenceKind.Cell;
            piece = FormatColumn(columnRelative, columnValue, anchorColumn, s) + FormatRow(rowRelative, rowValue, anchorRow, s);
        }
        else if (hasRow)
        {
            kind = ReferenceKind.WholeRow;
            piece = FormatRow(rowRelative, rowValue, anchorRow, s);
        }
        else
        {
            kind = ReferenceKind.WholeColumn;
            piece = FormatColumn(columnRelative, columnValue, anchorColumn, s);
        }

        return true;
    }

    private static bool TryParseAxis(string s, ref int j, out bool relative, out int value)
    {
        var n = s.Length;
        relative = true;
        value = 0;

        if (j < n && s[j] == '[')
        {
            j++;
            var negative = false;
            if (j < n && (s[j] == '-' || s[j] == '+'))
            {
                negative = s[j] == '-';
                j++;
            }

            if (!TryReadNumber(s, ref j, out var offset))
                return false;
            if (j >= n || s[j] != ']')
                return false;

            j++;
            value = negative ? -offset : offset;
            return true;
        }

        if (j < n && IsAsciiDigit(s[j]))
        {
            TryReadNumber(s, ref j, out value);
            relative = false;
            return true;
        }

        return true;
    }

    private static bool TryReadNumber(string s, ref int j, out int value)
    {
        var n = s.Length;
        var start = j;
        long acc = 0;
        while (j < n && IsAsciiDigit(s[j]))
        {
            acc = acc * 10 + (s[j] - '0');
            if (acc > SpreadsheetConstants.MaxNumberOfRows)
                acc = SpreadsheetConstants.MaxNumberOfRows + 1; // Cap to avoid overflow; will fail the bounds check.
            j++;
        }

        value = (int)acc;
        return j > start;
    }

    private static string FormatRow(bool relative, int value, int anchor, string formula)
    {
        var row = relative ? anchor + value : value;
        if (row is < 1 or > SpreadsheetConstants.MaxNumberOfRows)
            ThrowHelper.R1C1ReferenceOutOfBounds(formula);

        var rowNumber = row.ToString(NumberFormatInfo.InvariantInfo);
        return relative ? rowNumber : "$" + rowNumber;
    }

    private static string FormatColumn(bool relative, int value, int anchor, string formula)
    {
        var column = relative ? anchor + value : value;
        if (column is < 1 or > SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.R1C1ReferenceOutOfBounds(formula);

        var columnName = SpreadsheetUtility.GetColumnName(column);
        return relative ? columnName : "$" + columnName;
    }

    private static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c is '_' or '.';

    private static bool IsAsciiDigit(char c) => (uint)(c - '0') <= 9;
}
