using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;

public class ClassWithInvalidCellValueConverter
{
    [CellValueConverter(CellValueConverterType = typeof(ConverterWithoutParameterlessConstructor))]
    public string Property { get; set; } = null!;
}