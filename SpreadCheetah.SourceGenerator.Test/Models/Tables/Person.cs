using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Tables;

public class Person
{
    [ColumnHeader("First name")]
    public string? FirstName { get; set; }
    [ColumnHeader("Last name")]
    [ColumnOrder(1)]
    public string? LastName { get; set; }
    public int Age { get; set; }
}
