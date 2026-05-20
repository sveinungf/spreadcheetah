using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlWorksheetCell(IXLCell cell) : IWorksheetCell
{
    public int? IntValue => cell.GetValue<int?>();

    public decimal? DecimalValue => cell.GetValue<decimal?>();

    public string? StringValue => cell.Value.IsBlank ? null : cell.GetText();

    public DateTime? DateTimeValue => cell.TryGetValue<double>(out var value) ? DateTime.FromOADate(value) : null;

    public IStyle Style => new ClosedXmlStyle(cell.Style);

    public string? Formula => cell.FormulaA1;
}
