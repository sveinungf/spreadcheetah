using SpreadCheetah.Helpers;
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

    private readonly struct ParsedReference(int length, ReferenceKind kind, bool rowRelative, int row, bool columnRelative, int column)
    {
        public int Length { get; } = length;
        public ReferenceKind Kind { get; } = kind;
        public bool RowRelative { get; } = rowRelative;
        public int Row { get; } = row;
        public bool ColumnRelative { get; } = columnRelative;
        public int Column { get; } = column;
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
                && TryParseReference(formula, i, row, column, out var first))
            {
                var afterFirst = i + first.Length;
                var rightChar = afterFirst < n ? formula[afterFirst] : '\0';

                // Combine into a range when followed by ':' and another reference.
                if (rightChar == ':'
                    && TryParseReference(formula, afterFirst + 1, row, column, out var second))
                {
                    var afterSecond = afterFirst + 1 + second.Length;
                    var rightChar2 = afterSecond < n ? formula[afterSecond] : '\0';
                    if (!IsWordChar(rightChar2) && rightChar2 is not ('[' or '!'))
                    {
                        WriteReference(sb, first, selfDuplicate: false);
                        sb.Append(':');
                        WriteReference(sb, second, selfDuplicate: false);
                        i = afterSecond;
                        continue;
                    }
                }

                if (!IsWordChar(rightChar) && rightChar is not ('[' or '!'))
                {
                    WriteReference(sb, first, selfDuplicate: true);
                    i = afterFirst;
                    continue;
                }
            }

            sb.Append(c);
            i++;
        }

        return sb.ToString();
    }

    private static void WriteReference(StringBuilder sb, in ParsedReference reference, bool selfDuplicate)
    {
        switch (reference.Kind)
        {
            case ReferenceKind.Cell:
                WriteColumnPart(sb, reference);
                WriteRowPart(sb, reference);
                break;
            case ReferenceKind.WholeRow:
                // A whole row reference in the A1 notation has the form "5:5".
                WriteRowPart(sb, reference);
                if (selfDuplicate)
                {
                    sb.Append(':');
                    WriteRowPart(sb, reference);
                }

                break;
            default:
                // A whole column reference in the A1 notation has the form "C:C".
                WriteColumnPart(sb, reference);
                if (selfDuplicate)
                {
                    sb.Append(':');
                    WriteColumnPart(sb, reference);
                }

                break;
        }
    }

    private static void WriteRowPart(StringBuilder sb, in ParsedReference reference)
    {
        if (!reference.RowRelative)
            sb.Append('$');

        sb.Append(reference.Row);
    }

    private static void WriteColumnPart(StringBuilder sb, in ParsedReference reference)
    {
        if (!reference.ColumnRelative)
            sb.Append('$');

        sb.Append(SpreadsheetUtility.GetColumnName(reference.Column));
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

    private static bool TryParseReference(string s, int start, int anchorRow, int anchorColumn, out ParsedReference reference)
    {
        reference = default;

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

        var length = j - start;
        var kind = (hasRow, hasColumn) switch
        {
            (true, true) => ReferenceKind.Cell,
            (true, false) => ReferenceKind.WholeRow,
            _ => ReferenceKind.WholeColumn
        };

        var row = hasRow ? Resolve(rowRelative, rowValue, anchorRow, SpreadsheetConstants.MaxNumberOfRows, s) : 0;
        var column = hasColumn ? Resolve(columnRelative, columnValue, anchorColumn, SpreadsheetConstants.MaxNumberOfColumns, s) : 0;

        reference = new ParsedReference(length, kind, rowRelative, row, columnRelative, column);
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

    private static int Resolve(bool relative, int value, int anchor, int max, string formula)
    {
        var result = relative ? anchor + value : value;
        if (result < 1 || result > max)
            ThrowHelper.R1C1ReferenceOutOfBounds(formula);

        return result;
    }

    private static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c is '_' or '.';

    private static bool IsAsciiDigit(char c) => (uint)(c - '0') <= 9;
}
