using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

[InheritColumns]
[InferColumnHeaders(typeof(InferColumnHeadersResources), Prefix = "Header_")]
public class DerivedClassWithInfer : BaseClassWithoutInfer
{
    public required string Model { get; init; }
}
