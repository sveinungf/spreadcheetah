using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables;

/// <summary>
/// Provides column options for a table.
/// </summary>
public sealed class TableColumnOptions
{
    /// <summary>
    /// The total row label for the column.
    /// </summary>
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

    /// <summary>
    /// The total row function for the column.
    /// </summary>
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