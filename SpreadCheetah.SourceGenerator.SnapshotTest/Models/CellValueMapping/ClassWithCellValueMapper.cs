using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGeneration.Internal;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueMapping;

public class ClassWithCellValueMapper
{
    [ColumnHeader("Property")]
    [CellValueMapper(CellValueMapperType = typeof(StringValueMapper))]
    public string? Property { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueMapper(CellValueMapperType = typeof(IntValueMapper))]
    public int? Property1 { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueMapper(CellValueMapperType = typeof(DecimalValueMapper))]
    public decimal Property2 { get; set; }
}