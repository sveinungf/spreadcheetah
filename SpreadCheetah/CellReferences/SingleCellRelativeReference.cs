using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace SpreadCheetah.CellReferences;

internal readonly partial record struct SingleCellRelativeReference
{
    private const int MatchTimeoutMilliseconds = 1000;
    private const RegexOptions Options = RegexOptions.ExplicitCapture;

    /// <summary>Example: A1</summary>
    [StringSyntax(StringSyntaxAttribute.Regex)]
    private const string Pattern = "^(?<column>[A-Z]{1,3})(?<row>[1-9][0-9]{0,6})$";

#if NET7_0_OR_GREATER
    [GeneratedRegex(Pattern, Options, MatchTimeoutMilliseconds)]
    private static partial Regex Regex();
#else
    private static Regex RegexInstance { get; } = new(Pattern, Options, TimeSpan.FromMilliseconds(MatchTimeoutMilliseconds));
    private static Regex Regex() => RegexInstance;
#endif

    public string Reference { get; }

    /// <summary>Column 'A' becomes column number 1.</summary>
    public int Column { get; }

    /// <summary>Row number starts at 1.</summary>
    public int Row { get; }

    private SingleCellRelativeReference(string reference, int column, int row)
    {
        Reference = reference;
        Column = column;
        Row = row;
    }

    public static SingleCellRelativeReference Create(string value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        var column = 0;
        int row = 0;
        var match = Regex().Match(value);
        if (!match.Success ||
            match.Captures is not { Count: 2 } captures ||
            !TryParseColumnNumber(captures[0], out column) ||
            !TryParseInteger(captures[1], out row))
        {
            ThrowHelper.SingleCellReferenceInvalid(paramName);
        }

        return new SingleCellRelativeReference(value, column, row);
    }

    // TODO: Make optimized variant and move to SpreadsheetUtility
    private static bool TryParseColumnNumber(ReadOnlySpan<char> columnName, out int columnNumber)
    {
        columnNumber = 0;
        var pow = 1;
        for (int i = columnName.Length - 1; i >= 0; i--)
        {
            columnNumber += (columnName[i] - 'A' + 1) * pow;
            pow *= 26;
        }

        return true;
    }

    private static bool TryParseColumnNumber(Capture capture, out int columnNumber)
    {
#if NET6_0_OR_GREATER
        return TryParseColumnNumber(capture.ValueSpan, out columnNumber);
#else
        return TryParseColumnNumber(capture.Value.AsSpan(), out columnNumber);
#endif
    }

    private static bool TryParseInteger(Capture capture, out int result)
    {
#if NET6_0_OR_GREATER
        return int.TryParse(capture.ValueSpan, out result);
#else
        return int.TryParse(capture.Value, out result);
#endif
    }
}
