using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

public class ClassWithCellStyleOnTruncatedProperty
{
    [CellStyle("Name style")]
    [CellValueTruncate(10)]
    public required string Name { get; init; }
}
