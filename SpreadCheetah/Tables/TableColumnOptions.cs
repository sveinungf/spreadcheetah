using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

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
        get;
        set
        {
            if (value is not null && TotalRowFunction is not null)
                TableThrowHelper.ColumnCanNotHaveTotalRowFunctionAndLabel();

            field = value;
        }
    }

    /// <summary>
    /// The total row function for the column.
    /// </summary>
    public TableTotalRowFunction? TotalRowFunction
    {
        get;
        set
        {
            if (value is not null && TotalRowLabel is not null)
                TableThrowHelper.ColumnCanNotHaveTotalRowFunctionAndLabel();

            field = Guard.DefinedEnumValue(value);
        }
    }

    /// <summary>
    /// The style to apply to the cell in the total row for the column.
    /// </summary>
    public Style? TotalRowStyle { get; set; }

    internal bool AffectsTotalRow => this is not
    {
        TotalRowFunction: null,
        TotalRowLabel: null,
        TotalRowStyle: null
    };
}