namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides outline options to be used with <see cref="WorksheetOptions"/>.
/// </summary>
public sealed class OutlineOptions
{
    /// <summary>
    /// Flag indicating whether summary rows appear below detail in an outline, when applying an outline.
    /// <br/>
    /// When <see langword="true"/> a summary row is inserted below the detailed data being summarized and a new outline level is established on that row.
    /// <br/>
    /// When <see langword="false"/> a summary row is inserted above the detailed data being summarized and a new outline level is established on that row.
    /// </summary>
    public bool? SummaryBelow { get; set; }

    /// <summary>
    /// Flag indicating whether summary columns appear to the right of detail in an outline, when applying an outline.
    /// <br/>
    /// When <see langword="true"/> a summary column is inserted to the right of the detailed data being summarized and a new outline level is established on that column.
    /// <br/>
    /// When <see langword="false"/> a summary column is inserted to the left of the detailed data being summarized and a new outline level is established on that column.
    /// </summary>
    public bool? SummaryRight { get; set; }
}