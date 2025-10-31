using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

[InferColumnHeaders(typeof(InferColumnHeadersResources))]
public class ClassWithMultipleProperties
{
    public int Id { get; set; }
    public string? Name { get; set; }
    public decimal Price { get; set; }
}
