namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowTypeProperty(
    string Name,
    CellFormat? CellFormat,
    CellStyle? CellStyle,
    CellValueConverter? CellValueConverter,
    CellValueTruncate? CellValueTruncate,
    ColumnHeaderInfo? ColumnHeader,
    ColumnWidth? ColumnWidth,
    PropertyFormula? Formula)
{
    public bool HasStyle => CellFormat is not null || CellStyle is not null || Formula is { Type: FormulaType.HyperlinkFromUri };
}