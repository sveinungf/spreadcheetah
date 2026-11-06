using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides optional row options to be used when adding a row to the spreadsheet.
/// </summary>
public sealed class RowOptions
{
    /// <summary>
    /// Used on the summary row to indicate that the next outline level (detail rows) is currently in a collapsed state.
    /// </summary>
    public bool? Collapsed { get; set; }

    /// <summary>
    /// The default style for cells in the row. This style will be applied to all cells in the row
    /// unless they have their own style set.
    /// </summary>
    public Style? DefaultStyle { get; set; }

    /// <summary>
    /// The height of the row. Must be between 0 and 409. When not set Excel will default to 15.
    /// </summary>
    public double? Height { get; set => field = Guard.RowHeightInRange(value); }

    /// <summary>
    /// Set whether the row should be hidden or not.
    /// </summary>
    public bool? Hidden { get; set; }

    /// <summary>
    /// The outline level for row grouping. Must be between 0 and 7, where 0 means no grouping.
    /// Higher values indicate deeper nesting levels in the row hierarchy.
    /// </summary>
    public int? OutlineLevel { get; set => field = Guard.OutlineLevelInRange(value); }
}
