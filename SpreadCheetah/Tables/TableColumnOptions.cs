using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables;

public sealed class TableColumnOptions
{
    public string? TotalRowLabel { get; set; } // TODO: Validate in setter

    public TableTotalRowFunction? TotalRowFunction
    {
        get => _totalRowFunction;
        set => _totalRowFunction = Guard.DefinedEnumValue(value);
    }

    private TableTotalRowFunction? _totalRowFunction;

    internal bool AffectsTotalRow => TotalRowLabel is not null || TotalRowFunction is not null;
}