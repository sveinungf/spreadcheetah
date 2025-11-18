using ClosedXML.Excel;
using SpreadCheetah.Tables;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertTableColumn(IXLTableField tableField)
    : ISpreadsheetAssertTableColumn
{
    public string Name => tableField.Name;
    public string? TotalRowLabel => tableField.TotalsRowLabel;

    public TableTotalRowFunction? TotalRowFunction => tableField.TotalsRowFunction switch
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
