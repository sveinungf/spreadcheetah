using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models;

[InheritColumns(InheritedColumnOrder.InheritedColumnsLast)]
public record RecordClassWithInheritanceAndStartFromClassProperties(string Name, bool Value, int Age)
    : RecordClassWithSingleProperty(Value);