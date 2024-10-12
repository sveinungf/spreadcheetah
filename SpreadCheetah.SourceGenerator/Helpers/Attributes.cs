namespace SpreadCheetah.SourceGenerator.Helpers;

internal static class Attributes
{
    public const string CellStyle = "CellStyleAttribute";
    public const string CellValueConverter = "CellValueConverterAttribute";
    public const string CellValueConverterGeneric = "CellValueConverterAttribute`1";
    public const string CellValueTruncate = "CellValueTruncateAttribute";
    public const string ColumnHeader = "ColumnHeaderAttribute";
    public const string ColumnOrder = "ColumnOrderAttribute";
    public const string ColumnWidth = "ColumnWidthAttribute";

    // TODO: Get rid of these
    public const string CellValueTruncateFqn = "SpreadCheetah.SourceGeneration.CellValueTruncateAttribute";
    public const string ColumnOrderFqn = "SpreadCheetah.SourceGeneration.ColumnOrderAttribute";
    public const string InheritColumnsFqn = "SpreadCheetah.SourceGeneration.InheritColumnsAttribute";
    public const string GenerationOptionsFqn = "SpreadCheetah.SourceGeneration.WorksheetRowGenerationOptionsAttribute";
}
