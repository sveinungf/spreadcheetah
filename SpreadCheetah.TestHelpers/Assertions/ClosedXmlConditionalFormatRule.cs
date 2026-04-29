using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlConditionalFormatRule(IXLConditionalFormat conditionalFormat)
    : IConditionalFormatRule
{
    public string CellRangeReference => conditionalFormat.Range.ToString();

    public bool IsUniqueValuesRule => conditionalFormat.ConditionalFormatType is XLConditionalFormatType.IsUnique;

    public ISpreadsheetAssertStyle Style => new ClosedXmlAssertStyle(conditionalFormat.Style);
}
