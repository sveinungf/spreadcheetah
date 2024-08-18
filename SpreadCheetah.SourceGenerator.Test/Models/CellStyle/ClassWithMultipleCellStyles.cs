using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellStyle;

public class ClassWithMultipleCellStyles
{
    [CellStyle("Name")]
    public required string FirstName { get; set; }

    public string? MiddleName { get; set; }

    [CellStyle("Name")]
    public required string LastName { get; set; }

    [CellStyle("Age")]
    public required int Age { get; set; }
}