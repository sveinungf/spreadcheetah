namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models;

public class ClassWithMultipleProperties
{
    public string FirstName { get; set; } = "";
    public string? MiddleName { get; set; }
    public string LastName { get; set; } = "";
    public int Age { get; set; }
    public bool Employed { get; set; }
    public double Score { get; set; }
}
