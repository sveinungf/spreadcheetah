using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnOrdering;

public class ClassWithColumnOrderForAllProperties
{
    [ColumnOrder(2)]
    public string FirstName { get; set; } = "";

    [ColumnOrder(3)]
    public string? MiddleName { get; set; }

    [ColumnOrder(1)]
    public string LastName { get; set; } = "";

    [ColumnOrder(5)]
    public int Age { get; set; }

    [ColumnOrder(4)]
    public bool Employed { get; set; }

    [ColumnOrder(6)]
    public double Score { get; set; }
}
