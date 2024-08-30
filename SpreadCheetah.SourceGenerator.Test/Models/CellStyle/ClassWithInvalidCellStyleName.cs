using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

public class ClassWithInvalidCellStyleName
{
    [CellStyle("")]
    public string? Name { get; set; }
}
