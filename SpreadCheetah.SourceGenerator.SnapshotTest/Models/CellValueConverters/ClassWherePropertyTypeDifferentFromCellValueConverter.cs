using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;

public class ClassWherePropertyTypeDifferentFromCellValueConverter
{
    [CellValueConverter(CellValueConverterType = typeof(NullableIntValueConverter))]
    public string Property { get; set; } = null!;
    
    [CellValueConverter(CellValueConverterType = typeof(NullableIntValueConverter))]
    public int? Property1 { get; set; }
}