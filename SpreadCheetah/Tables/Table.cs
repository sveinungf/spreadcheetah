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
        Style = style; // TODO: Validate
        Name = name; // TODO: Validate
    }

    internal Dictionary<int, TableColumnOptions>? ColumnOptions { get; private set; }

    public TableColumnOptions Column(int columnNumber)
    {
        if (columnNumber is < 1 or > SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        // TODO: Is there a limit of the number of columns in a table?
        // TODO: Consider storing the max columnNumber in a property, for easy lookup later
        ColumnOptions ??= [];
        if (!ColumnOptions.TryGetValue(columnNumber, out var options))
        {
            options = new TableColumnOptions();
            ColumnOptions[columnNumber] = options;
        }

        return options;
    }
}
