using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellFormat;

public class ClassWithCellCustomFormat
{
    [CellFormat("#.0#")]
    public decimal Price { get; set; }
}
