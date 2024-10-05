using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InheritColumns;

[InheritColumns(DefaultColumnOrder = InheritedColumnsOrder.InheritedColumnsLast)]
public class ClassWithInheritedColumnsLast : BaseClass
{
    public string? DerivedClassProperty { get; set; }
}
