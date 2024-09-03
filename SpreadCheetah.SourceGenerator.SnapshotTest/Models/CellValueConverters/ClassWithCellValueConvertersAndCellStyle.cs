using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;

public class ClassWithCellValueConvertersAndCellStyle
{
    [ColumnHeader("Property")]
    [CellValueConverter(typeof(StringValueConverter))]
    [CellStyle("Test")]
    public string? Property { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueConverter(typeof(NullableIntValueConverter))]
    public int? Property1 { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueConverter(typeof(DecimalValueConverter))]
    public decimal Property2 { get; set; }
}