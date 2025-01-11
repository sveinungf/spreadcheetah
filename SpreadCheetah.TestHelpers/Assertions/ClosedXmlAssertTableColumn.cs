using ClosedXML.Excel;
using SpreadCheetah.Tables;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertTableColumn(IXLTableField field)
    : ISpreadsheetAssertTableColumn
{
    public string Name => field.Name;
    public string? TotalRowLabel => field.TotalsRowLabel;

    public TableTotalRowFunction? TotalRowFunction => field.TotalsRowFunction switch
    {
        XLTotalsRowFunction.Average => TableTotalRowFunction.Average,
        XLTotalsRowFunction.Count => TableTotalRowFunction.Count,
        XLTotalsRowFunction.CountNumbers => TableTotalRowFunction.CountNumbers,
        XLTotalsRowFunction.Maximum => TableTotalRowFunction.Maximum,
        XLTotalsRowFunction.Minimum => TableTotalRowFunction.Minimum,
        XLTotalsRowFunction.StandardDeviation => TableTotalRowFunction.StandardDeviation,
        XLTotalsRowFunction.Sum => TableTotalRowFunction.Sum,
        XLTotalsRowFunction.Variance => TableTotalRowFunction.Variance,
        _ => null
    };
}
