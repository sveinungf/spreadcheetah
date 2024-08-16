using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnStyle;

public class ClassWithColumnStyle
{
    [ColumnStyle("Price style")]
    public decimal Price { get; set; }
}
