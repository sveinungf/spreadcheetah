using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns]
public record RecordClass2LevelOfInheritanceInheritedColumnsFirstParentIgnoreInheritance(
    string ClassProperty,
    string Name,
    bool Value,
    int Age)
    : RecordClassWithIgnoreInheritance(Name, Value, Age);