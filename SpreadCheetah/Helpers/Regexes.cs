using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace SpreadCheetah.Helpers;

internal static partial class Regexes
{
    private const int TimeoutMillis = 1000;
    private const string ColumnName = "[A-Z]{1,3}";
    private const string RowNumber = "[1-9][0-9]{0,6}";

    /// <summary>
    /// Absolute (e.g. <c>$A$1</c>) or relative (e.g. <c>A1</c>) column reference.
    /// </summary>
    private const string ColumnReferencePattern = @"^\$?" + ColumnName;

    /// <summary>
    /// Absolute (e.g. <c>$1</c>) or relative (e.g. <c>1</c>) row reference.
    /// </summary>
    private const string RowReferencePattern = @"^\$?" + RowNumber;

    /// <summary>
    /// Optional absolute (e.g. <c>:$B$2</c>) or relative (e.g. <c>:B2</c>) range reference.
    /// </summary>
    private const string OptionalRangeReferencePattern = @"^(?::\$?" + ColumnName + @"\$?" + RowNumber + ")?$";

    private const string TableNameCellReferencePattern = "^" + ColumnName + "[0-9]{1,7}";

    private const string TableNameValidCharactersPattern = @"^[A-Z_\\][A-Z0-9._\\]*$";

    private const RegexOptions TableNameCellReferenceOptions = RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture;
    private const RegexOptions TableNameValidCharactersOptions = RegexOptions.IgnoreCase;

#if NET9_0_OR_GREATER
    [GeneratedRegex(ColumnReferencePattern, RegexOptions.None, TimeoutMillis)]
    public static partial Regex ColumnReference { get; }

    [GeneratedRegex(RowReferencePattern, RegexOptions.None, TimeoutMillis)]
    public static partial Regex RowReference { get; }

    [GeneratedRegex(OptionalRangeReferencePattern, RegexOptions.None, TimeoutMillis)]
    public static partial Regex OptionalRangeReference { get; }

    [GeneratedRegex(TableNameCellReferencePattern, TableNameCellReferenceOptions, TimeoutMillis)]
    public static partial Regex TableNameCellReference { get; }

    [GeneratedRegex(TableNameValidCharactersPattern, TableNameValidCharactersOptions, TimeoutMillis)]
    public static partial Regex TableNameValidCharacters { get; }
#else
    private static TimeSpan Timeout => TimeSpan.FromMilliseconds(TimeoutMillis);

    public static Regex ColumnReference { get; } = new(ColumnReferencePattern, RegexOptions.None, Timeout);
    public static Regex RowReference { get; } = new(RowReferencePattern, RegexOptions.None, Timeout);
    public static Regex OptionalRangeReference { get; } = new(OptionalRangeReferencePattern, RegexOptions.None, Timeout);
    public static Regex TableNameCellReference { get; } = new(TableNameCellReferencePattern, TableNameCellReferenceOptions, Timeout);
    public static Regex TableNameValidCharacters { get; } = new(TableNameValidCharactersPattern, TableNameValidCharactersOptions, Timeout);
#endif
}
