using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides column options to be used with <see cref="WorksheetOptions"/>.
/// </summary>
public sealed class ColumnOptions
{
    /// <summary>
    /// The default style for cells in the column. This style will be applied to all cells in the column
    /// unless they have their own style set, or if the row they are in has a style set.
    /// </summary>
    public Style? DefaultStyle { get; set; }

    /// <summary>
    /// The width of the column. The number represents how many characters can be displayed in the standard font.
    /// It must be between 0 and 255. When not set Excel will default to approximately 8.89.
    /// </summary>
    public double? Width { get; set => field = Guard.ColumnWidthInRange(value); }

    /// <summary>
    /// Set whether the column should be hidden or not. Defaults to <see langword="false"/>.
    /// </summary>
    public bool Hidden { get; set; }

    /// <summary>
    /// The outline level for column grouping. Must be between 0 and 7, where 0 means no grouping.
    /// Higher values indicate deeper nesting levels in the column hierarchy.
    /// </summary>
    public int? OutlineLevel { get; set => field = Guard.OutlineLevelInRange(value); }

    /// <summary>
    /// Used on the summary column to indicate that the next outline level (detail columns) is currently in a collapsed state.
    /// </summary>
    public bool? Collapsed { get; set; }
}
