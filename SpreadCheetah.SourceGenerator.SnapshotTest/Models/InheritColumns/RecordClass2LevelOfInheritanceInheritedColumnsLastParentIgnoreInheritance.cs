using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;

[InheritColumns(InheritedColumnOrder = InheritedColumnOrder.InheritedColumnsLast)]
public record RecordClass2LevelOfInheritanceInheritedColumnsLastParentIgnoreInheritance(
    string ClassProperty,
    string Name,
    bool Value,
    int Age)
    : RecordClassWithIgnoreInheritance(Name, Value, Age);