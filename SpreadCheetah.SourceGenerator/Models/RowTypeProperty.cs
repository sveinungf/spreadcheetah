namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowTypeProperty(
    string Name,
    ColumnHeaderInfo? ColumnHeader,
    ColumnStyle? ColumnStyle,
    ColumnWidth? ColumnWidth,
    CellValueTruncate? CellValueTruncate);