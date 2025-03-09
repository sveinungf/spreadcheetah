using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables;

/// <summary>
/// Represents an Excel table.
/// </summary>
public sealed class Table
{
    internal TableStyle Style { get; }
    internal string? Name { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="Table"/> class.
    /// The name parameter is optional. If not set, then a name will be generated automatically,
    /// starting with <c>"Table1"</c> for the first added table. Otherwise the name must:
    /// <list type="bullet">
    ///   <item>Contain only letters, numbers, periods and underscores.</item>
    ///   <item>Start with a letter, an underscore, or a backslash.</item>
    ///   <item>Not exceed 255 characters.</item>
    ///   <item>Not be equal to <c>"C"</c>, <c>"c"</c>, <c>"R"</c>, or <c>"r"</c>.</item>
    ///   <item>Not be a cell reference, such as <c>"A1"</c> or <c>"R1C1"</c>.</item>
    ///   <item>Be unique.</item>
    /// </list>
    /// </summary>
    public Table(TableStyle style, string? name = null)
    {
        Style = Guard.DefinedEnumValue(style);

        if (name is not null)
            EnsureValidTableName(name);

        Name = name;
    }

    /// <summary>Alternative shading or banding in columns.</summary>
    public bool BandedColumns { get; set; }

    /// <summary>Alternative shading or banding in rows.</summary>
    public bool BandedRows { get; set; } = true;

    /// <summary>
    /// Explicitly set the number of columns in the table.
    /// The number of columns is normally deduced from the header row column.
    /// This option should only be used when the number of columns should be different from the header row.
    /// </summary>
    public int? NumberOfColumns
    {
        get => _numberOfColumns;
        set
        {
            if (value <= 0)
                throw new ArgumentOutOfRangeException(nameof(value), value, "Value must be greater than 0.");

            if (value < _maxColumnNumber)
                TableThrowHelper.NumberOfColumnsLessThanGreatestColumnNumber(value);

            _numberOfColumns = value;
        }
    }

    private int? _numberOfColumns;

    private static void EnsureValidTableName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            ThrowHelper.NameEmptyOrWhiteSpace(nameof(name));

        if (name.Length > 255)
            ThrowHelper.NameTooLong(255, nameof(name));

        if (name is "C" or "c" or "R" or "r")
            TableThrowHelper.NameCanNotBeCorR(nameof(name));

        if (!Regexes.TableNameValidCharacters().IsMatch(name))
            TableThrowHelper.NameHasInvalidCharacters(nameof(name));

        if (Regexes.TableNameCellReference().IsMatch(name))
            TableThrowHelper.NameIsCellReference(nameof(name));
    }

    private int _maxColumnNumber;
    internal Dictionary<int, TableColumnOptions>? ColumnOptions { get; private set; }

    /// <summary>
    /// Get options for a column in the table. The first column has column number 1.
    /// </summary>
    public TableColumnOptions Column(int columnNumber)
    {
        if (columnNumber is < 1 or > SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        if (NumberOfColumns is { } columns && columnNumber > columns)
            TableThrowHelper.ColumnNumberGreaterThanNumberOfColumns(columnNumber);

        ColumnOptions ??= [];
        if (!ColumnOptions.TryGetValue(columnNumber, out var options))
        {
            options = new TableColumnOptions();
            ColumnOptions[columnNumber] = options;
            _maxColumnNumber = Math.Max(_maxColumnNumber, columnNumber);
        }

        return options;
    }
}
