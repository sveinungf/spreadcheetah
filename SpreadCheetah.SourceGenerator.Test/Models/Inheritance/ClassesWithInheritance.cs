using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Inheritance;

public class ClassAnimal
{
    public DateTime DateOfBirth { get; set; }
}

[InheritColumns(DefaultColumnOrder = InheritedColumnsOrder.InheritedColumnsLast)]
public class ClassMammal : ClassAnimal
{
    public bool CanWalk { get; set; }
}

[InheritColumns]
public class ClassDog : ClassMammal
{
    public string Breed { get; set; } = "";
}