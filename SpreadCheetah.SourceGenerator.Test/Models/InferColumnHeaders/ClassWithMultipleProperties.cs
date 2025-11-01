using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

[InferColumnHeaders(typeof(InferColumnHeadersResources))]
public class ClassWithMultipleProperties
{
    public int Id { get; set; }

#pragma warning disable SPCH1010 // Missing property for ColumnHeader
    public string? Name { get; set; }
#pragma warning restore SPCH1010 // Missing property for ColumnHeader

    [ColumnHeader("The price")]
    public decimal Price { get; set; }
}
