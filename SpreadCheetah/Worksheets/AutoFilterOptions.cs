using SpreadCheetah.Helpers;
using System.Text.RegularExpressions;

namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides auto filter options to be used with <see cref="WorksheetOptions"/>.
/// </summary>
public class AutoFilterOptions
{
    private static Regex Regex { get; } = new Regex("^[A-Z]{1,3}[0-9]{1,7}:[A-Z]{1,3}[0-9]{1,7}$", RegexOptions.None, TimeSpan.FromSeconds(1));

    /// <summary>
    /// The cell range to filter. Not exposed publicly in case there will be a more dynamic way to set the range in the future.
    /// </summary>
    internal string CellRange { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="AutoFilterOptions"/> class for a range of cells.
    /// The cell range must be in the A1 reference style. Some examples:
    /// <list type="bullet">
    ///   <item><term><c>A1:A4</c></term><description>References the range from cell A1 to A4.</description></item>
    ///   <item><term><c>A1:E5</c></term><description>References the range from cell A1 to E5.</description></item>
    ///   <item><term><c>A1:A1048576</c></term><description>References all cells in column A.</description></item>
    ///   <item><term><c>A5:XFD5</c></term><description>References all cells in row 5.</description></item>
    /// </list>
    /// </summary>
    public AutoFilterOptions(string cellRange)
    {
        if (!Regex.IsMatch(cellRange))
            ThrowHelper.CellReferenceInvalid(nameof(cellRange));

        CellRange = cellRange;
    }
}