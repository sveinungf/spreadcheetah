using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

[InheritColumns]
public class DerivedClassWithInferFromBaseClass : BaseClassWithInfer
{
    public required string Model { get; init; }
}
