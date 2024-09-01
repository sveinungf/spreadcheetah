using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

public class ClassWithCellStyle
{
    [CellStyle("Price style")]
    public decimal Price { get; set; }
}
