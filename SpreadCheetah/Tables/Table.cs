using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables;

public sealed class Table
{
    internal TableStyle Style { get; }
    internal string? Name { get; }

    public bool BandedColumns { get; set; }
    public bool BandedRows { get; set; } = true;
    public int? NumberOfColumns { get; set; } // TODO: Validate in setter. 0 should not be allowed. For comment: Should only be used if number of columns is different from header row columns.

    public Table(TableStyle style, string? name = null)
    {
        Style = Guard.DefinedEnumValue(style);

        if (name is not null)
            EnsureValidTableName(name);

        Name = name;
    }

    private static void EnsureValidTableName(string name)
    {
        if (name.Length == 0)
            throw new NotImplementedException(); // TODO: Empty string not allowed.
        if (name.Length > 255)
            throw new NotImplementedException(); // TODO: A table name can have up to 255 characters.
        if (name is "C" or "c" or "R" or "r")
            throw new NotImplementedException(); // TODO: You can't use "C", "c", "R", or "r" for the name.

        // TODO: Name must start with a letter, underscore, or '\'.
        // TODO: Use letters, numbers, periods, and underscore characters for the rest of the name (and also '\')
        if (!Regexes.TableNameValidCharacters().IsMatch(name))
            throw new NotImplementedException();

        // TODO: Don't allow A1 or R1C1 references (also with different casing)
        if (Regexes.TableNameCellReference().IsMatch(name))
            throw new NotImplementedException();
    }

    internal Dictionary<int, TableColumnOptions>? ColumnOptions { get; private set; }

    public TableColumnOptions Column(int columnNumber)
    {
        if (columnNumber is < 1 or > SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        // TODO: Is there a limit of the number of columns in a table?
        ColumnOptions ??= [];
        if (!ColumnOptions.TryGetValue(columnNumber, out var options))
        {
            options = new TableColumnOptions();
            ColumnOptions[columnNumber] = options;
        }

        return options;
    }
}
