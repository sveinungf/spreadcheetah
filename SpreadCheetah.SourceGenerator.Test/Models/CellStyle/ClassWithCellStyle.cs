using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

public class ClassWithCellStyle
{
    [CellStyle("Price style")]
    public double Price { get; set; }
}
