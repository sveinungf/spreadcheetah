using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns(InheritedColumnOrder = InheritedColumnOrder.InheritedColumnsLast)]
public record RecordClass2LevelOfInheritanceInheritedColumnsLastParentInheritColumnsFirst(
    string OwnValue,
    bool Value,
    string ClassValue) :
    RecordClassWithInheritedColumnsLast(Value, ClassValue);