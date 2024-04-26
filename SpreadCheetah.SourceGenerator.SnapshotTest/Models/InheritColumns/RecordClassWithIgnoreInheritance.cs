namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

public record RecordClassWithIgnoreInheritance(string Name, bool Value, int Age)
    : RecordClassWithSingleProperty(Value);