using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns(InheritedColumnOrder.InheritedColumnsLast)]
public record RecordClassWithInheritedColumnsLast(bool Value, string ClassValue) : RecordClassWithSingleProperty(Value);