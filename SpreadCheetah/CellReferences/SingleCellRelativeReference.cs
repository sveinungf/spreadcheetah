using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
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
            match.Groups is not { Count: 3 } groups ||
            !TryParseColumnName(groups[1], out column) ||
            !TryParseInteger(groups[2], out row))
        {
            ThrowHelper.SingleCellReferenceInvalid(paramName);
        }

        return new SingleCellRelativeReference(value, column, row);
    }

    private static bool TryParseColumnName(Capture capture, out int columnNumber)
    {
#if NET6_0_OR_GREATER
        return SpreadsheetUtility.TryParseColumnName(capture.ValueSpan, out columnNumber);
#else
        return SpreadsheetUtility.TryParseColumnName(capture.Value.AsSpan(), out columnNumber);
#endif
    }

    private static bool TryParseInteger(Capture capture, out int result)
    {
#if NET6_0_OR_GREATER
        return int.TryParse(capture.ValueSpan, NumberStyles.Integer, NumberFormatInfo.InvariantInfo, out result);
#else
        return int.TryParse(capture.Value, NumberStyles.Integer, NumberFormatInfo.InvariantInfo, out result);
#endif
    }
}
