using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;

public class ClassWithCellValueConvertersAndCellStyle
{
    [ColumnHeader("Property")]
    [CellValueConverter(CellValueConverterType = typeof(StringValueConverter))]
    [CellStyle("Test")]
    public string? Property { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueConverter(CellValueConverterType = typeof(NullableIntValueConverter))]
    public int? Property1 { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueConverter(CellValueConverterType = typeof(DecimalValueConverter))]
    public decimal Property2 { get; set; }
}