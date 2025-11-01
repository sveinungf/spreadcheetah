using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

[InferColumnHeaders(typeof(InferColumnHeadersResources), Prefix = "Header_")]
public abstract class BaseClassWithInfer
{
    public required string Make { get; init; }
}