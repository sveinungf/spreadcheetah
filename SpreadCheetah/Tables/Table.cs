using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables;

public sealed class Table
{
    internal string Name { get; }
    internal TableStyle Style { get; }

    public bool BandedColumns { get; set; }
    public bool BandedRows { get; set; } = true;
    public int? NumberOfColumns { get; set; } // TODO: Validate in setter. Should only be used if number of columns is different from header row columns.

    public Table(string name, TableStyle style)
    {
        Name = name; // TODO: Validate
        Style = style; // TODO: Validate
    }

    // TODO: Make copy in immutable type. Use e.g. List for now.
    internal SortedDictionary<int, TableColumnOptions>? ColumnOptions { get; private set; }
    internal bool HasTotalRow() => ColumnOptions?.Values.Any(x => x.AffectsTotalRow) ?? false;

    public TableColumnOptions Column(int columnNumber)
    {
        if (columnNumber is < 1 or > SpreadsheetConstants.MaxNumberOfColumns)
            ThrowHelper.ColumnNumberInvalid(nameof(columnNumber), columnNumber);

        // TODO: Is there a limit of the number of columns in a table?
        // TODO: Consider having AutoFilter stuff in TableColumnOptions
        ColumnOptions ??= [];
        if (!ColumnOptions.TryGetValue(columnNumber, out var options))
        {
            options = new TableColumnOptions();
            ColumnOptions[columnNumber] = options;
        }

        return options;
    }
}
