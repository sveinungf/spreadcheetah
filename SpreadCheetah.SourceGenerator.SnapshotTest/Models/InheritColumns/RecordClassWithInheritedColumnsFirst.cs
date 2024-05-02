using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns]
public record RecordClassWithInheritedColumnsFirst(bool Value, string ClassProperty)
    : RecordClassWithSingleProperty(Value);