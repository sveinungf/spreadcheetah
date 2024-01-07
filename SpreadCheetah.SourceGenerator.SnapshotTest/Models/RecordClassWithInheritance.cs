namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models;

public record RecordClassWithInheritance(string Name, bool Value, int Age)
    : RecordClassWithSingleProperty(Value);