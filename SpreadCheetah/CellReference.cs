using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace SpreadCheetah;

internal readonly partial record struct CellReference
{
    private const int MatchTimeoutMilliseconds = 1000;

    /// <summary>Example: A1:B10</summary>
    [StringSyntax(StringSyntaxAttribute.Regex)]
    private const string RelativeCellRangePattern = "^[A-Z]{1,3}[0-9]{1,7}:[A-Z]{1,3}[0-9]{1,7}$";

    /// <summary>
    /// Examples:
    /// <list type="bullet">
    ///   <item><term><c>A1</c></term> <description>Cell A1, relative reference.</description></item>
    ///   <item><term><c>$C$4</c></term> <description>Cell C4, absolute reference.</description></item>
    ///   <item><term><c>$D6</c></term> <description>Cell D6, mixed reference.</description></item>
    ///   <item><term><c>A1:E5</c></term><description>Cell range A1 to E5, relative references.</description></item>
    ///   <item><term><c>$C$4:$H$10</c></term><description>Cell range C4 to H10, absolute references.</description></item>
    /// </list>
    /// </summary>
    [StringSyntax(StringSyntaxAttribute.Regex)]
    private const string RelativeOrAbsoluteCellOrCellRangePattern = @"^\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?$";

#if !NET7_0_OR_GREATER
    private static Regex RelativeCellRangeRegexInstance { get; } = new(RelativeCellRangePattern, RegexOptions.None, TimeSpan.FromMilliseconds(MatchTimeoutMilliseconds));
    private static Regex RelativeCellRangeRegex() => RelativeCellRangeRegexInstance;

    private static Regex RelativeOrAbsoluteCellOrCellRangeRegexInstance { get; } = new(RelativeOrAbsoluteCellOrCellRangePattern, RegexOptions.None, TimeSpan.FromMilliseconds(MatchTimeoutMilliseconds));
    private static Regex RelativeOrAbsoluteCellOrCellRangeRegex() => RelativeOrAbsoluteCellOrCellRangeRegexInstance;
#else
    [GeneratedRegex(RelativeCellRangePattern, RegexOptions.None, MatchTimeoutMilliseconds)]
    private static partial Regex RelativeCellRangeRegex();

    [GeneratedRegex(RelativeOrAbsoluteCellOrCellRangePattern, RegexOptions.None, MatchTimeoutMilliseconds)]
    private static partial Regex RelativeOrAbsoluteCellOrCellRangeRegex();
#endif

    public string Reference { get; }

    private CellReference(string reference) => Reference = reference;

    private static bool TryCreate(string value, bool allowSingleCell, CellReferenceType type, [NotNullWhen(true)] out CellReference? reference)
    {
        reference = null;

        if (string.IsNullOrWhiteSpace(value))
            return false;

        var match = type switch
        {
            CellReferenceType.Relative when !allowSingleCell => RelativeCellRangeRegex().IsMatch(value),
            CellReferenceType.RelativeOrAbsolute when allowSingleCell => RelativeOrAbsoluteCellOrCellRangeRegex().IsMatch(value),
            _ => false
        };

        if (match)
            reference = new CellReference(value);

        return match;
    }

    public static CellReference Create(string value, bool allowSingleCell, CellReferenceType type, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        if (!TryCreate(value, allowSingleCell, type, out var reference))
        {
            if (allowSingleCell)
                ThrowHelper.CellOrCellRangeReferenceInvalid(paramName);
            else
                ThrowHelper.CellRangeReferenceInvalid(paramName);
        }

        return reference.Value;
    }
}
