using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.DefaultColumnWidth;

[DefaultColumnWidth(10)]
public class ClassWithDefaultColumnWidth
{
    public int Id { get; set; }

    [ColumnWidth(20)]
    public string? Name { get; set; }
}
