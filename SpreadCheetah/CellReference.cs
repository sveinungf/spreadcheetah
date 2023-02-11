using SpreadCheetah.Helpers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace SpreadCheetah;

internal readonly record struct CellReference
{
    /// <summary>Example: A1:B10</summary>
    private static Regex RelativeCellRangeRegex { get; } = new("^[A-Z]{1,3}[0-9]{1,7}:[A-Z]{1,3}[0-9]{1,7}$", RegexOptions.None, TimeSpan.FromSeconds(1));

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
    private static Regex RelativeOrAbsoluteCellOrCellRangeRegex { get; } = new(@"^\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?$", RegexOptions.None, TimeSpan.FromSeconds(1));

    public string Reference { get; }

    private CellReference(string reference) => Reference = reference;

    private static bool TryCreate(string value, bool allowSingleCell, CellReferenceType type, [NotNullWhen(true)] out CellReference? reference)
    {
        reference = null;

        if (string.IsNullOrWhiteSpace(value))
            return false;

        var match = type switch
        {
            CellReferenceType.Relative when !allowSingleCell => RelativeCellRangeRegex.IsMatch(value),
            CellReferenceType.RelativeOrAbsolute when allowSingleCell => RelativeOrAbsoluteCellOrCellRangeRegex.IsMatch(value),
            _ => false
        };

        if (match)
            reference = new CellReference(value);

        return match;
    }

    public static CellReference Create(string value, bool allowSingleCell, CellReferenceType type, [CallerArgumentExpression("value")] string? paramName = null)
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
