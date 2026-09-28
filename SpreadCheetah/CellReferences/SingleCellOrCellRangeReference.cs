using SpreadCheetah.Helpers;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace SpreadCheetah.CellReferences;

internal readonly record struct SingleCellOrCellRangeReference
{
    public string Reference { get; }

    /// <summary>Column of the top-left cell; column 'A' becomes column number 1.</summary>
    public ushort Column { get; }

    /// <summary>Row of the top-left cell; row number starts at 1.</summary>
    public uint Row { get; }

    private SingleCellOrCellRangeReference(string reference, ushort column, uint row)
    {
        Reference = reference;
        Column = column;
        Row = row;
    }

    public static SingleCellOrCellRangeReference Create(string value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        var valueSpan = value.AsSpan();

        var columnLength = GetMatchLength(Regexes.ColumnReference, valueSpan, paramName);
        var columnSpan = valueSpan[..columnLength];
        if (columnSpan[0] == '$')
            columnSpan = columnSpan[1..];
        if (!SpreadsheetUtility.TryParseColumnName(columnSpan, out var columnNumber))
            ThrowHelper.SingleCellReferenceInvalid(paramName);

        var rowLength = GetMatchLength(Regexes.RowReference, valueSpan[columnLength..], paramName);
        var rowSpan = valueSpan.Slice(columnLength, rowLength);
        if (rowSpan[0] == '$')
            rowSpan = rowSpan[1..];

        if (!uint.TryParse(rowSpan, NumberStyles.None, CultureInfo.InvariantCulture, out var row))
            ThrowHelper.SingleCellOrCellRangeReferenceInvalid(paramName);

        if (!Regexes.OptionalRangeReference.IsMatch(valueSpan[(columnLength + rowLength)..]))
            ThrowHelper.SingleCellOrCellRangeReferenceInvalid(paramName);

        return new SingleCellOrCellRangeReference(value, (ushort)columnNumber, row);
    }

#if NET7_0_OR_GREATER
    private static int GetMatchLength(Regex regex, ReadOnlySpan<char> span, string? paramName)
    {
        var enumerator = regex.EnumerateMatches(span);
        if (!enumerator.MoveNext())
            ThrowHelper.SingleCellReferenceInvalid(paramName);

        return enumerator.Current.Length;
    }
#else
    private static int GetMatchLength(Regex regex, ReadOnlySpan<char> span, string? paramName)
    {
        var match = regex.Match(span.ToString());
        if (!match.Success)
            ThrowHelper.SingleCellReferenceInvalid(paramName);

        return match.Length;
    }
#endif
}
