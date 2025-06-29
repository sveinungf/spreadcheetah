using SpreadCheetah.Helpers;

namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides column options to be used with <see cref="WorksheetOptions"/>.
/// </summary>
public class ColumnOptions
{
    /// <summary>
    /// The width of the column. The number represents how many characters can be displayed in the standard font.
    /// It must be between 0 and 255. When not set Excel will default to approximately 8.89.
    /// </summary>
    public double? Width
    {
        get => _width;
        set => _width = Guard.ColumnWidthInRange(value);
    }

    private double? _width;

    /// <summary>
    /// Is the column hidden or not. Defaults to <see langword="false"/>.
    /// </summary>
    public bool Hidden { get; set; }
}
