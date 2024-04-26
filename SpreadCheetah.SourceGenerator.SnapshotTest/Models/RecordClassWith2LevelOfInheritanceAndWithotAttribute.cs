
namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models;

public record RecordClassWith2LevelOfInheritanceAndWithoutAttribute(string FirstName, string Name, bool Value, int Age)
    : RecordClassWithInheritance(Name, Value, Age);