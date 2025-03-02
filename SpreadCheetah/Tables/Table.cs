using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables;

public sealed class Table
{
    internal TableStyle Style { get; }
    internal string? Name { get; }

    public bool BandedColumns { get; set; }
    public bool BandedRows { get; set; } = true;

    // TODO: For comment: Should only be used if number of columns is different from header row columns.
    public int? NumberOfColumns
    {
        get => _numberOfColumns;
        set => _numberOfColumns = value <= 0
            ? throw new ArgumentOutOfRangeException(nameof(value), value, "Value must be greater than 0.")
            : value;
    }

    private int? _numberOfColumns;

    public Table(TableStyle style, string? name = null)
    {
        Style = Guard.DefinedEnumValue(style);

        if (name is not null)
            EnsureValidTableName(name);

        Name = name;
    }

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

    internal Dictionary<int, TableColumnOptions>? ColumnOptions { get; private set; }

    public TableColumnOptions Column(int columnNumber)
    {
        if (columnNumber is < 1 or > SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        ColumnOptions ??= [];
        if (!ColumnOptions.TryGetValue(columnNumber, out var options))
        {
            options = new TableColumnOptions();
            ColumnOptions[columnNumber] = options;
        }

        return options;
    }
}
