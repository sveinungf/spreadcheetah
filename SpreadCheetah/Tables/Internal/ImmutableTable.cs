namespace SpreadCheetah.Tables.Internal;

internal readonly record struct ImmutableTable(
    string Name,
    TableStyle Style,
    bool BandedColumns,
    bool BandedRows,
    int? NumberOfColumns,
    int? TotalRowMaxColumnNumber,
    IReadOnlyDictionary<int, ImmutableTableColumnOptions>? ColumnOptions)
{
    public bool HasTotalRow => TotalRowMaxColumnNumber is not null;

    public static ImmutableTable From(Table table, string tableName)
    {
        Dictionary<int, ImmutableTableColumnOptions>? columnOptions = null;
        int? totalRowMaxColumnNumber = null;

        if (table.ColumnOptions is not null)
        {
            columnOptions = new(table.ColumnOptions.Count);

            foreach (var (key, value) in table.ColumnOptions)
            {
                // TODO: Ignore column if key > table.NumberOfColumns?
                columnOptions[key] = ImmutableTableColumnOptions.From(value);

                if (value.AffectsTotalRow)
                {
                    var currentMax = totalRowMaxColumnNumber ?? 0;
                    if (key > currentMax)
                        totalRowMaxColumnNumber = key;
                }
            }
        }

        return new ImmutableTable(
            Name: tableName,
            Style: table.Style,
            BandedColumns: table.BandedColumns,
            BandedRows: table.BandedRows,
            NumberOfColumns: table.NumberOfColumns,
            TotalRowMaxColumnNumber: totalRowMaxColumnNumber,
            ColumnOptions: columnOptions);
    }
}
