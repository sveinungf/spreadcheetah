using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;

public class ConverterWithoutParameterlessConstructor : CellValueConverter<string>
{
    private readonly string? _test;

    private ConverterWithoutParameterlessConstructor()
    {
    }

    public ConverterWithoutParameterlessConstructor(string test)
    {
        _test = test;
    }

    public override DataCell ConvertToCell(string value)
    {
        return new DataCell(_test);
    }
}