using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InheritColumns;

[InheritColumns]
public class TwoLevelInheritanceClassWithInheritedColumns : ClassWithInheritedColumnsLast
{
    public string? LeafClassProperty { get; set; }
}
