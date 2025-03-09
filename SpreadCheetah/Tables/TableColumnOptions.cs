using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables;

public sealed class TableColumnOptions
{
    public string? TotalRowLabel
    {
        get => _totalRowLabel;
        set
        {
            if (value is not null && TotalRowFunction is not null)
                TableThrowHelper.ColumnCanNotHaveTotalRowFunctionAndLabel();

            _totalRowLabel = value;
        }
    }

    private string? _totalRowLabel;

    public TableTotalRowFunction? TotalRowFunction
    {
        get => _totalRowFunction;
        set
        {
            if (value is not null && TotalRowLabel is not null)
                TableThrowHelper.ColumnCanNotHaveTotalRowFunctionAndLabel();

            _totalRowFunction = Guard.DefinedEnumValue(value);
        }
    }

    private TableTotalRowFunction? _totalRowFunction;

    internal bool AffectsTotalRow => TotalRowLabel is not null || TotalRowFunction is not null;
}