using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InheritColumns;

[InheritColumns]
public class ClassWithInheritedColumnsFirst : BaseClass
{
    public string? DerivedClassProperty { get; set; }
}
