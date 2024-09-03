using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;

public class ClassWithCellValueConverters
{
    [ColumnHeader("Property")]
    [CellValueConverter(typeof(StringValueConverter))]
    public string? Property { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueConverter(typeof(NullableIntValueConverter))]
    public int? Property1 { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueConverter(typeof(DecimalValueConverter))]
    public decimal Property2 { get; set; }
}