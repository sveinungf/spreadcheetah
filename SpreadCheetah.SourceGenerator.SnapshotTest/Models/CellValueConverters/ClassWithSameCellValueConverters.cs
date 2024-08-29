using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;

public class ClassWithSameCellValueConverters
{
    [ColumnHeader("Property")]
    [CellValueConverter(CellValueConverterType = typeof(StringValueConverter))]
    public string? Property { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueConverter(CellValueConverterType = typeof(StringValueConverter))]
    public string? Property1 { get; set; }
    
    [ColumnHeader("Property1")]
    [CellValueConverter(CellValueConverterType = typeof(DecimalValueConverter))]
    public decimal Property2 { get; set; }
}