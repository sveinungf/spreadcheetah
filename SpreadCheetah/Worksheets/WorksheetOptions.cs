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
    public double? DefaultColumnWidth { get; set => field = Guard.ColumnWidthInRange(value); }

    /// <summary>
    /// The number of left-most columns that should be frozen.
    /// </summary>
    public int? FrozenColumns { get; set => field = Guard.FrozenColumnsInRange(value); }

    /// <summary>
    /// The number of top-most rows that should be frozen.
    /// </summary>
    public int? FrozenRows { get; set => field = Guard.FrozenRowsInRange(value); }

    /// <summary>
    /// Option to hide a worksheet.
    /// </summary>
    public WorksheetVisibility Visibility { get; set => field = Guard.DefinedEnumValue(value); }

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
