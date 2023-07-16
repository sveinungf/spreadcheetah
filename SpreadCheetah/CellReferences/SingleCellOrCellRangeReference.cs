using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace SpreadCheetah.CellReferences;

internal readonly partial record struct SingleCellOrCellRangeReference
{
    private const int MatchTimeoutMilliseconds = 1000;
    private const RegexOptions Options = RegexOptions.None;

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
    private const string Pattern = @"^\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6})?$";

#if NET7_0_OR_GREATER
    [GeneratedRegex(Pattern, Options, MatchTimeoutMilliseconds)]
    private static partial Regex Regex();
#else
    private static Regex RegexInstance { get; } = new(Pattern, Options, TimeSpan.FromMilliseconds(MatchTimeoutMilliseconds));
    private static Regex Regex() => RegexInstance;
#endif

    public string Reference { get; }

    private SingleCellOrCellRangeReference(string reference) => Reference = reference;

    public static SingleCellOrCellRangeReference Create(string value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        if (!Regex().IsMatch(value))
            ThrowHelper.SingleCellOrCellRangeReferenceInvalid(paramName);

        return new SingleCellOrCellRangeReference(value);
    }
}
