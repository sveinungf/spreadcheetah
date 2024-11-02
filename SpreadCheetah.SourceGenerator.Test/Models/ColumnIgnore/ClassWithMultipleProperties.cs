using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnIgnore;

public class ClassWithMultipleProperties
{
    [ColumnIgnore]
    public int Id { get; set; }

    public string? Name { get; set; }

    [ColumnIgnore]
    public decimal Price { get; set; }
}
