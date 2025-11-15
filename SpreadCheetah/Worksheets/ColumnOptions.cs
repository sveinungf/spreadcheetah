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
    /// Is the column hidden or not. Defaults to <see langword="false"/>.
    /// </summary>
    public bool Hidden { get; set; }
}
