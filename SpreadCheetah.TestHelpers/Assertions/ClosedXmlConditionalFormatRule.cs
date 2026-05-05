using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlConditionalFormatRule(IXLConditionalFormat conditionalFormat)
    : IConditionalFormatRule
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

    public bool IsUniqueValuesRule => conditionalFormat.ConditionalFormatType is XLConditionalFormatType.IsUnique;

    public ISpreadsheetAssertStyle Style => new ClosedXmlAssertStyle(conditionalFormat.Style);
}
