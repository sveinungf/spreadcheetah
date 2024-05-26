namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowTypeProperty(
    string Name,
    ColumnHeaderInfo? ColumnHeader,
    CellValueTruncate? CellValueTruncate);