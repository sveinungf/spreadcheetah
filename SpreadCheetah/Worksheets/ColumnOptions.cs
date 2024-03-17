namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides column options to be used with <see cref="WorksheetOptions"/>.
/// </summary>
public class ColumnOptions
{
    /// <summary>
    /// The width of the column. The number represents how many characters can be displayed in the standard font.
    /// Must be between 0 and 255. When not set Excel will default to 8.43.
    /// </summary>
    public double? Width
    {
        get => _width;
        set => _width = value is <= 0 or > 255
            ? throw new ArgumentOutOfRangeException(nameof(value), value, "Column width must be between 0 and 255.")
            : value;
    }

    private double? _width;

    /// <summary>
    /// Is the column hidden or not.  Defaults to false.
    /// </summary>
    public bool Hidden { get; set; }
}
