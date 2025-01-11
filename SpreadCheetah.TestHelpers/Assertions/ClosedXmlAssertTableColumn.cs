using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertTableColumn(IXLTableField field)
    : ISpreadsheetAssertTableColumn
{
    public string Name => field.Name;
}
