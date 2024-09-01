namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowTypeProperty(
    string Name,
    ColumnHeaderInfo? ColumnHeader,
    CellStyle? CellStyle,
    ColumnWidth? ColumnWidth,
    CellValueTruncate? CellValueTruncate);