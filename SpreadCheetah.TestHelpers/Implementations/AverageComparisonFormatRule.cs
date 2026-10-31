using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;
using SpreadCheetah.TestHelpers.Models;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class AverageComparisonFormatRule(
    IXLConditionalFormat conditionalFormat,
    AverageComparisonOperator op)
    : IAverageComparisonFormatRule
{
    public string CellRangeReference
    {
        get
        {
            var rangeAddress = conditionalFormat.Range.RangeAddress;
            return rangeAddress.NumberOfCells == 1
                ? rangeAddress.FirstAddress.ToStringRelative()
                : rangeAddress.ToStringRelative();
        }
    }

    public string Formula => throw new InvalidOperationException("The conditional format rule is not a formula rule.");
    public double NumericValue => throw new InvalidOperationException("The conditional format rule is not a cellIs rule.");
    public IReadOnlyList<double> NumericValues => throw new InvalidOperationException("The conditional format rule is not a cellIs rule.");
    public ConditionalFormatOperator RuleOperator => throw new InvalidOperationException("The conditional format rule is not a cellIs rule.");
    public bool IsCellIsRule => false;
    public bool IsDuplicateValuesRule => false;
    public bool IsMatchesFormulaRule => false;
    public bool IsUniqueValuesRule => false;
    public IStyle Style => ClosedXmlStyle.Create(conditionalFormat.Style);

    public AverageComparisonOperator AverageComparisonOperator => op;
}
