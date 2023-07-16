using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace SpreadCheetah.CellReferences;

internal readonly partial record struct CellRangeRelativeReference
{
    private const int MatchTimeoutMilliseconds = 1000;
    private const RegexOptions Options = RegexOptions.None;

    /// <summary>Example: A1:B10</summary>
    [StringSyntax(StringSyntaxAttribute.Regex)]
    private const string Pattern = "^[A-Z]{1,3}[1-9][0-9]{0,6}:[A-Z]{1,3}[1-9][0-9]{0,6}$";

#if NET7_0_OR_GREATER
    [GeneratedRegex(Pattern, Options, MatchTimeoutMilliseconds)]
    private static partial Regex Regex();
#else
    private static Regex RegexInstance { get; } = new(Pattern, Options, TimeSpan.FromMilliseconds(MatchTimeoutMilliseconds));
    private static Regex Regex() => RegexInstance;
#endif

    public string Reference { get; }

    private CellRangeRelativeReference(string reference) => Reference = reference;

    public static CellRangeRelativeReference Create(string value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        if (!Regex().IsMatch(value))
            ThrowHelper.CellRangeReferenceInvalid(paramName);

        return new CellRangeRelativeReference(value);
    }
}
