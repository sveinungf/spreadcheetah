using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.ColumnOrdering;

public record RecordClassWithColumnOrdering(
    [property: ColumnOrder(2)] string FirstName,
    [property: ColumnOrder(1)] string LastName,
    decimal Gpa,
    [property: ColumnOrder(3)] int Age);