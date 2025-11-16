using SpreadCheetah.Helpers;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace SpreadCheetah.CellReferences;

[StructLayout(LayoutKind.Auto)]
internal readonly partial record struct SingleCellRelativeReference
{
    private const int MatchTimeoutMilliseconds = 1000;

#if NET7_0_OR_GREATER
    [GeneratedRegex("^[A-Z]{1,3}", RegexOptions.None, MatchTimeoutMilliseconds)]
    private static partial Regex ColumnRegex();

    [GeneratedRegex("^[1-9][0-9]{0,6}$", RegexOptions.None, MatchTimeoutMilliseconds)]
    private static partial Regex RowRegex();
#else
    private static Regex Regex { get; } = new("^(?<column>[A-Z]{1,3})(?<row>[1-9][0-9]{0,6})$", RegexOptions.ExplicitCapture, TimeSpan.FromMilliseconds(MatchTimeoutMilliseconds));
#endif

    /// <summary>Column 'A' becomes column number 1.</summary>
    public ushort Column { get; }

    /// <summary>Row number starts at 1.</summary>
    public uint Row { get; }

    private SingleCellRelativeReference(ushort column, uint row)
    {
        Column = column;
        Row = row;
    }

#if NET7_0_OR_GREATER
    public static SingleCellRelativeReference Create(ReadOnlySpan<char> value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
#else
    public static SingleCellRelativeReference Create(ReadOnlySpan<char> value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
        => CreateInternal(value.ToString(), paramName ?? nameof(value));
    public static SingleCellRelativeReference Create(string value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
        => CreateInternal(value, paramName ?? nameof(value));

    private static SingleCellRelativeReference CreateInternal(string value, string paramName)
#endif
    {
        if (!TryParseColumnRow(value, out var column, out var row))
            ThrowHelper.SingleCellReferenceInvalid(paramName);

        return new SingleCellRelativeReference(column, row);
    }

#if NET7_0_OR_GREATER
    private static bool TryParseColumnRow(ReadOnlySpan<char> value, out ushort column, out uint row)
    {
        column = 0;
        row = 0;

        var enumerator = ColumnRegex().EnumerateMatches(value);
        if (!enumerator.MoveNext())
            return false;

        var columnNameLength = enumerator.Current.Length;
        if (!TryParseColumnName(value[..columnNameLength], out column))
            return false;

        var rowSpan = value[columnNameLength..];
        if (!RowRegex().IsMatch(rowSpan))
            return false;

        return uint.TryParse(rowSpan, NumberStyles.None, NumberFormatInfo.InvariantInfo, out row);
    }
#else
    private static bool TryParseColumnRow(string value, out ushort column, out uint row)
    {
        column = 0;
        row = 0;

        var match = Regex.Match(value);
        if (!match.Success || match.Groups is not [_, var columnGroup, var rowGroup])
            return false;

        if (!TryParseColumnName(columnGroup.Value, out column))
            return false;

        return uint.TryParse(rowGroup.Value, NumberStyles.None, NumberFormatInfo.InvariantInfo, out row);
    }
#endif

    private static bool TryParseColumnName(ReadOnlySpan<char> columnName, out ushort column)
    {
        if (!SpreadsheetUtility.TryParseColumnName(columnName, out var columnInt))
        {
            column = 0;
            return false;
        }

        column = (ushort)columnInt;
        return true;
    }
}
