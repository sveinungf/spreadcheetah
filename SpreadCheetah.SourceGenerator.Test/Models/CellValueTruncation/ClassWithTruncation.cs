using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueTruncation;

public class ClassWithTruncation
{
    [CellValueTruncate(15)]
    public string? Value { get; set; }
}
