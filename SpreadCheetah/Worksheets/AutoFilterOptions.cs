namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides auto filter options to be used with <see cref="WorksheetOptions"/>.
/// </summary>
public class AutoFilterOptions
{
    /// <summary>
    /// The cell range to filter. Not exposed publicly in case there will be a more dynamic way to set the range in the future.
    /// </summary>
    internal CellReference CellRange { get; }

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
        CellRange = CellReference.Create(cellRange, false, CellReferenceType.Relative);
    }
}