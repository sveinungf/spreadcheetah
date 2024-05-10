using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns(InheritedColumnOrder = InheritedColumnOrder.InheritedColumnsLast)]
public record RecordClass2LevelOfInheritanceInheritedColumnsLast(string OwnProperty, bool Value, string ClassProperty)
    : RecordClassWithInheritedColumnsFirst(Value, ClassProperty);