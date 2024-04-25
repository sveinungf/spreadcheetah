using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models;

[InheritedColumnOrdering(InheritedColumnsOrderingStrategy.StartFromInheritedProperties)]
public record RecordClassWithInheritanceAndStartFromInheritedProperties(string Name, bool Value, int Age)
    : RecordClassWithSingleProperty(Value);