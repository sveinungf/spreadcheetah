namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides optional row options to be used when adding a row to the spreadsheet.
/// </summary>
public class RowOptions
{
    /// <summary>
    /// The height of the row. Must be between 0 and 409. When not set Excel will default to 15.
    /// </summary>
    public double? Height
    {
        get => _height;
        set => _height = value is <= 0 or > 409
            ? throw new ArgumentOutOfRangeException(nameof(value), value, "Row height must be between 0 and 409.")
            : value;
    }

    private double? _height;
}
