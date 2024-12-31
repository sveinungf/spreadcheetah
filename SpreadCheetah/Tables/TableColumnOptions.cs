namespace SpreadCheetah.Tables;

public sealed class TableColumnOptions
{
    public string? TotalRowLabel { get; set; } // TODO: Validate in setter
    public TableTotalRowFunction? TotalRowFunction { get; set; } // TODO: Validate in setter

    internal bool AffectsTotalRow => TotalRowLabel is not null || TotalRowFunction is not null;
}