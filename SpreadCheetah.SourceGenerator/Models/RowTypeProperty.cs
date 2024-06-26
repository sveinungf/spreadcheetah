namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowTypeProperty(
    string Name,
    ColumnHeaderInfo? ColumnHeader,
    ColumnWidth? ColumnWidth,
    CellValueTruncate? CellValueTruncate);