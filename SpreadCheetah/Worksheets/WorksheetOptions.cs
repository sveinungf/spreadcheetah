using SpreadCheetah.Helpers;

namespace SpreadCheetah.Worksheets;

/// <summary>
/// Provides options to be used when starting a worksheet with <see cref="Spreadsheet"/>.
/// </summary>
public sealed class WorksheetOptions
{
    /// <summary>
    /// The default width of the columns in the worksheet. The number represents how many characters can be displayed in the standard font.
    /// It must be between 0 and 255. When not set Excel will default to approximately 8.89.
    /// </summary>
    public double? DefaultColumnWidth
    {
        get => _defaultColumnWidth;
        set => _defaultColumnWidth = Guard.ColumnWidthInRange(value);
    }

    private double? _defaultColumnWidth;

    /// <summary>
    /// The number of left-most columns that should be frozen.
    /// </summary>
    public int? FrozenColumns
    {
        get => _frozenColumns;
        set => _frozenColumns = value is < 1 or > SpreadsheetConstants.MaxNumberOfColumns
            ? throw new ArgumentOutOfRangeException(nameof(value), value, $"Number of frozen columns must be between 1 and {SpreadsheetConstants.MaxNumberOfColumns}")
            : value;
    }

    private int? _frozenColumns;

    /// <summary>
    /// The number of top-most rows that should be frozen.
    /// </summary>
    public int? FrozenRows
    {
        get => _frozenRows;
        set => _frozenRows = value is < 1 or > SpreadsheetConstants.MaxNumberOfRows
            ? throw new ArgumentOutOfRangeException(nameof(value), value, $"Number of frozen rows must be between 1 and {SpreadsheetConstants.MaxNumberOfRows}")
            : value;
    }

    private int? _frozenRows;

    /// <summary>
    /// Option to hide a worksheet.
    /// </summary>
    public WorksheetVisibility Visibility
    {
        get => _visibility;
        set => _visibility = Guard.DefinedEnumValue(value);
    }

    private WorksheetVisibility _visibility;

    /// <summary>
    /// Auto filtering options.
    /// </summary>
    public AutoFilterOptions? AutoFilter { get; set; }

    /// <summary>
    /// Option to show or hide the sheet grid.
    /// </summary>
    public bool? ShowGridlines { get; set; }

    internal SortedDictionary<int, ColumnOptions>? ColumnOptions { get; private set; }

    /// <summary>
    /// Get options for a column in the worksheet. The first column has column number 1.
    /// </summary>
    public ColumnOptions Column(int columnNumber)
    {
        if (columnNumber is < 1 or > SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        ColumnOptions ??= [];
        if (!ColumnOptions.TryGetValue(columnNumber, out var options))
        {
            options = new ColumnOptions();
            ColumnOptions[columnNumber] = options;
        }

        return options;
    }
}
