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
    private const string TableNameCellReferencePattern = "^([A-Z]{1,3}[0-9]{1,7}$|R[0-9]{1,7})";
    private const RegexOptions TableNameCellReferenceOptions = RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture;

#if NET7_0_OR_GREATER
    [GeneratedRegex(TableNameValidCharactersPattern, TableNameValidCharactersOptions, TimeoutMillis)]
    public static partial Regex TableNameValidCharacters();

    [GeneratedRegex(TableNameCellReferencePattern, TableNameCellReferenceOptions, TimeoutMillis)]
    public static partial Regex TableNameCellReference();
#else
    private static TimeSpan Timeout => TimeSpan.FromMilliseconds(TimeoutMillis);

    private static Regex TableNameValidCharactersInstance { get; } = new(TableNameValidCharactersPattern, TableNameValidCharactersOptions, Timeout);
    private static Regex TableNameCellReferenceInstance { get; } = new(TableNameCellReferencePattern, TableNameCellReferenceOptions, Timeout);
    public static Regex TableNameValidCharacters() => TableNameValidCharactersInstance;
    public static Regex TableNameCellReference() => TableNameCellReferenceInstance;
#endif
}
