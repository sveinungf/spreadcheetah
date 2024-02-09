using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnHeader;

public class ClassWithColumnHeaderForAllProperties
{
    [ColumnHeader("First name")]
    public string FirstName { get; set; } = "";

    [ColumnHeader("Middle name")]
    public string? MiddleName { get; set; }

    [ColumnHeader("Last name")]
    public string LastName { get; set; } = "";

    [ColumnHeader("Age")]
    public int Age { get; set; }

    [ColumnHeader("Employed (yes/no)")]
    public bool Employed { get; set; }

    [ColumnHeader("Score (decimal)")]
    public double Score { get; set; }
}
