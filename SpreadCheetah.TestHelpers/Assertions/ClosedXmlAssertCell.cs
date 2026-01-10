using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertCell(IXLCell cell) : ISpreadsheetAssertCell
{
    public int? IntValue => cell.GetValue<int?>();

    public decimal? DecimalValue => cell.GetValue<decimal?>();

    public string? StringValue => cell.Value.IsBlank ? null : cell.GetText();

    public DateTime? DateTimeValue => cell.TryGetValue<double>(out var value) ? DateTime.FromOADate(value) : null;

    public ISpreadsheetAssertStyle Style => new ClosedXmlAssertStyle(cell.Style);

    public string? Formula => cell.FormulaA1;
}
