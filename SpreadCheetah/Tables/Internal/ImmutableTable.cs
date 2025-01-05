namespace SpreadCheetah.Tables.Internal;

internal readonly record struct ImmutableTable(
    string? Name,
    TableStyle Style,
    bool BandedColumns,
    bool BandedRows,
    int? NumberOfColumns,
    IReadOnlyDictionary<int, ImmutableTableColumnOptions>? ColumnOptions)
{
    public static ImmutableTable From(Table table)
    {
        return new ImmutableTable(
            Name: table.Name,
            Style: table.Style,
            BandedColumns: table.BandedColumns,
            BandedRows: table.BandedRows,
            NumberOfColumns: table.NumberOfColumns,
            ColumnOptions: table.ColumnOptions?.ToDictionary(x => x.Key, x => ImmutableTableColumnOptions.From(x.Value)));
    }
}
