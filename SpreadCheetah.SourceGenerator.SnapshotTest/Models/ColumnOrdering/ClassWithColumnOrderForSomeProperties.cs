using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnOrdering;

public class ClassWithColumnOrderForSomeProperties
{
    public string FirstName { get; set; } = "";

    [ColumnOrder(-1000)]
    public string? MiddleName { get; set; }

    public string LastName { get; set; } = "";

    [ColumnOrder(500)]
    public int Age { get; set; }

    public bool Employed { get; set; }

    [ColumnOrder(2)]
    public double Score { get; set; }
}
