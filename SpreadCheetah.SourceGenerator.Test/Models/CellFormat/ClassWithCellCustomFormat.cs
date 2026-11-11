using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellFormat;

public class ClassWithCellCustomFormat
{
    [CellFormat("#.0#")]
    public double Price { get; set; }
}
