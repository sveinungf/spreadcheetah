using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlConditionalFormatRule(
    IXLConditionalFormat conditionalFormat)
    : IConditionalFormatRule
{
    public bool IsDuplicateValuesRule => conditionalFormat.ConditionalFormatType is XLConditionalFormatType.IsDuplicate;
    public bool IsMatchesFormulaRule => conditionalFormat.ConditionalFormatType is XLConditionalFormatType.Expression;
    public bool IsUniqueValuesRule => conditionalFormat.ConditionalFormatType is XLConditionalFormatType.IsUnique;
    public IStyle Style => ClosedXmlStyle.Create(conditionalFormat.Style);

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

    public string Formula
    {
        get
        {
            if (!IsMatchesFormulaRule)
                throw new InvalidOperationException("The conditional format rule is not a formula rule.");

            var value = conditionalFormat.Values.Values.SingleOrDefault();
            if (value is not { IsFormula: true, Value.Length: > 0 })
                throw new InvalidOperationException("The conditional format rule does not have a formula value.");

            return value.Value;
        }
    }
}
