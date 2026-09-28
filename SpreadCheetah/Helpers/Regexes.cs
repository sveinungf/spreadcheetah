using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace SpreadCheetah.Helpers;

internal static partial class Regexes
{
    private const int TimeoutMillis = 1000;

    [StringSyntax(StringSyntaxAttribute.Regex)]
    private const string TableNameValidCharactersPattern = @"^[A-Z_\\][A-Z0-9._\\]*$";
    private const RegexOptions TableNameValidCharactersOptions = RegexOptions.IgnoreCase;

    [StringSyntax(StringSyntaxAttribute.Regex)]
    private const string TableNameCellReferencePattern = "^[A-Z]{1,3}[0-9]{1,7}";
    private const RegexOptions TableNameCellReferenceOptions = RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture;

#if NET9_0_OR_GREATER
    [GeneratedRegex(TableNameValidCharactersPattern, TableNameValidCharactersOptions, TimeoutMillis)]
    public static partial Regex TableNameValidCharacters { get; }

    [GeneratedRegex(TableNameCellReferencePattern, TableNameCellReferenceOptions, TimeoutMillis)]
    public static partial Regex TableNameCellReference { get; }
#else
    private static TimeSpan Timeout => TimeSpan.FromMilliseconds(TimeoutMillis);

    public static Regex TableNameValidCharacters { get; } = new(TableNameValidCharactersPattern, TableNameValidCharactersOptions, Timeout);
    public static Regex TableNameCellReference { get; } = new(TableNameCellReferencePattern, TableNameCellReferenceOptions, Timeout);
#endif
}
