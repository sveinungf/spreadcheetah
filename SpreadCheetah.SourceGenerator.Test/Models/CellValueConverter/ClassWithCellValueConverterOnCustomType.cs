using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;

public class ClassWithCellValueConverterOnCustomType
{
    public string Property { get; init; } = null!;

    [CellValueConverter(typeof(NullToDashValueConverter<object?>))]
    public object? ComplexProperty { get; init; }

    [CellStyle("PercentType")]
    [CellValueConverter(typeof(PercentToNumberConverter))]
    public Percent? PercentType { get; init; }
}