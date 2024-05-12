using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Inheritance;

public record RecordClassAnimal(DateTime DateOfBirth);

[InheritColumns(DefaultColumnOrder = InheritedColumnsOrder.InheritedColumnsLast)]
public record RecordClassMammal(bool CanWalk, DateTime DateOfBirth)
    : RecordClassAnimal(DateOfBirth);

[InheritColumns]
public record RecordClassDog(string Breed, bool CanWalk, DateTime DateOfBirth)
    : RecordClassMammal(CanWalk, DateOfBirth);
