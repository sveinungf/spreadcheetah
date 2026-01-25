using SpreadCheetah.Styling.Internal;

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

    public static ImmutableTable From(Table table, string tableName, StyleManager styleManager)
    {
        Dictionary<int, ImmutableTableColumnOptions>? columnOptions = null;
        int? totalRowMaxColumnNumber = null;

        if (table.ColumnOptions is not null)
        {
            columnOptions = new(table.ColumnOptions.Count);

            foreach (var (columnNo, options) in table.ColumnOptions)
            {
                columnOptions[columnNo] = ImmutableTableColumnOptions.From(options, styleManager);

                if (options.AffectsTotalRow)
                {
                    var currentMax = totalRowMaxColumnNumber ?? 0;
                    if (columnNo > currentMax)
                        totalRowMaxColumnNumber = columnNo;
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
