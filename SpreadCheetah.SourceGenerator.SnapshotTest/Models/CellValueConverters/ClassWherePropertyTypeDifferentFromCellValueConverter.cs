using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;

public class ClassWherePropertyTypeDifferentFromCellValueConverter
{
    [CellValueConverter(typeof(NullableIntValueConverter))]
    public string Property { get; set; } = null!;
    
    [CellValueConverter(typeof(ClassWithSingleProperty))]
    public int? Property1 { get; set; }
}