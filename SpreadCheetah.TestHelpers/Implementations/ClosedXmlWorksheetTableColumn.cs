using ClosedXML.Excel;
using SpreadCheetah.Tables;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlWorksheetTableColumn(IXLTableField tableField)
    : IWorksheetTableColumn
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
