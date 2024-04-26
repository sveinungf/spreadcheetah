using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models;

[InheritColumns]
public record RecordClassWithInheritanceAndStartFromInheritedProperties(string Name, bool Value, int Age)
    : RecordClassWithSingleProperty(Value);