using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides optional row options to be used when adding a row to the spreadsheet.
/// </summary>
public sealed class RowOptions
{
    /// <summary>
    /// The default style for cells in the row. This style will be applied to all cells in the row
    /// unless they have their own style set.
    /// </summary>
    public Style? DefaultStyle { get; set; }

    /// <summary>
    /// The height of the row. Must be between 0 and 409. When not set Excel will default to 15.
    /// </summary>
    public double? Height { get; set => field = Guard.RowHeightInRange(value); }
}
