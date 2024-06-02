using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnWidth;

public class ClassWithColumnWidth
{
    [ColumnWidth(20)]
    public string? Name { get; set; }
}
