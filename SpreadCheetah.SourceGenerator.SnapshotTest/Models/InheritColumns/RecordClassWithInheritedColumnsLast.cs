using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns(DefaultColumnOrder = InheritedColumnOrder.InheritedColumnsLast)]
public record RecordClassWithInheritedColumnsLast(bool Value, string ClassValue) : RecordClassWithSingleProperty(Value);