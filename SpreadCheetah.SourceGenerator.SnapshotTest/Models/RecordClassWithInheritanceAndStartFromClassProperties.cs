using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models;

[InheritedColumnOrdering(InheritedColumnsOrderingStrategy.StartFromClassProperties)]
public record RecordClassWithInheritanceAndStartFromClassProperties(string Name, bool Value, int Age)
    : RecordClassWithSingleProperty(Value);