using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

[InferColumnHeaders(typeof(InferColumnHeadersResources))]
public class ClassWithMultipleProperties
{
    public int Id { get; set; } // Can't use property 'X' on 'Y' for column header, because the property does not have a public getter. If 'Y' is a resource file, make sure it has its access modifier set to public.
    public string? Name { get; set; } // 'Y' does not have a property with name 'X' to use for column header.
    public decimal Price { get; set; }
}
