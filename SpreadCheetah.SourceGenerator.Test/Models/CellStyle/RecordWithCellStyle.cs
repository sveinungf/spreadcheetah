using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

public record RecordWithCellStyle
{
    [CellStyle("Name style")]
    public required string Name { get; init; }
}
