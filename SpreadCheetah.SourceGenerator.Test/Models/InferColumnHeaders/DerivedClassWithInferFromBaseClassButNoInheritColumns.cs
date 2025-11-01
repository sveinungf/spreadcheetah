namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

public class DerivedClassWithInferFromBaseClassButNoInheritColumns : BaseClassWithInfer
{
    public required string Model { get; init; }
}
