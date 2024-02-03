using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertCell(IXLCell cell) : ISpreadsheetAssertCell
{
    public int? IntValue => cell.GetValue<int>();

    public decimal? DecimalValue => cell.GetValue<decimal>();

    public string? StringValue => cell.GetText();
}
