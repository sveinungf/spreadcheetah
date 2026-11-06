namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowTypeProperty
{
    public required string Name { get; init; }
    public required CellFormat? CellFormat { get; init; }
    public required CellStyle? CellStyle { get; init; }
    public required CellValueConverter? CellValueConverter { get; init; }
    public required CellValueTruncate? CellValueTruncate { get; init; }
    public required ColumnHeaderInfo? ColumnHeader { get; init; }
    public required ColumnWidth? ColumnWidth { get; init; }
    public required PropertyFormula? Formula { get; init; }

    public bool HasStyle => CellFormat is not null || CellStyle is not null || Formula is { Type: FormulaType.HyperlinkFromUri };
}