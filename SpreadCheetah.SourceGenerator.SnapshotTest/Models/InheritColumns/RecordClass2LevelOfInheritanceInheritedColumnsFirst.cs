using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns]
public record RecordClass2LevelOfInheritanceInheritedColumnsFirst(string OwnProperty, bool Value, string ClassProperty)
    : RecordClassWithInheritedColumnsFirst(Value, ClassProperty);