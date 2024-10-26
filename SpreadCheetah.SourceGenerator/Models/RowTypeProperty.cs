namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowTypeProperty(
    string Name,
    CellFormat? CellFormat,
    CellStyle? CellStyle,
    CellValueConverter? CellValueConverter,
    CellValueTruncate? CellValueTruncate,
    ColumnHeaderInfo? ColumnHeader,
    ColumnWidth? ColumnWidth)
{
    public bool HasStyle => this is not { CellFormat: null, CellStyle: null };
}