namespace SpreadCheetah.Tables.Internal;

internal readonly record struct ImmutableTable(
    string Name,
    TableStyle Style,
    bool BandedColumns,
    bool BandedRows,
    bool HasTotalRow,
    int? NumberOfColumns,
    IReadOnlyDictionary<int, ImmutableTableColumnOptions>? ColumnOptions)
{
    public static ImmutableTable From(Table table, string tableName)
    {
        Dictionary<int, ImmutableTableColumnOptions>? columnOptions = null;
        var hasTotalRow = false;

        if (table.ColumnOptions is not null)
        {
            columnOptions = new(table.ColumnOptions.Count);

            foreach (var (key, value) in table.ColumnOptions)
            {
                columnOptions[key] = ImmutableTableColumnOptions.From(value);
                hasTotalRow = hasTotalRow || value.AffectsTotalRow;
            }
        }

        return new ImmutableTable(
            Name: tableName,
            Style: table.Style,
            BandedColumns: table.BandedColumns,
            BandedRows: table.BandedRows,
            HasTotalRow: hasTotalRow,
            NumberOfColumns: table.NumberOfColumns,
            ColumnOptions: columnOptions);
    }
}
