using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns]
public record RecordClass2LevelOfInheritanceInheritedColumnsFirst(
    string ClassProperty,
    string Name,
    bool Value,
    int Age)
    : RecordClassWithIgnoreInheritance(Name, Value, Age);