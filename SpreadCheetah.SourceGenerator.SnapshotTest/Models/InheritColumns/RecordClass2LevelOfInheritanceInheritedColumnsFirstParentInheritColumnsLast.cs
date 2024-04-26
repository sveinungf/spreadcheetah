using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns]
public record RecordClass2LevelOfInheritanceInheritedColumnsFirstParentInheritColumnsLast(
    string OwnValue,
    bool Value,
    string ClassValue) :
    RecordClassWithInheritedColumnsLast(Value, ClassValue);