namespace SpreadCheetah.Tables.Internal;

internal readonly record struct ImmutableTableColumnOptions(
    string? TotalRowLabel,
    TableTotalRowFunction? TotalRowFunction)
{
    public static ImmutableTableColumnOptions From(TableColumnOptions options)
    {
        return new ImmutableTableColumnOptions(options.TotalRowLabel, options.TotalRowFunction);
    }
}