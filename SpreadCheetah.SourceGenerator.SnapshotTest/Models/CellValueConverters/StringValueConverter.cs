using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;

internal class StringValueConverter : CellValueConverter<string>
{
    public override DataCell ConvertToCell(string value)
    {
        return new DataCell(value);
    }
}

internal class IntValueConverter : CellValueConverter<int?>
{
    public override DataCell ConvertToCell(int? value)
    {
        return new DataCell(value);
    }
}

internal class DecimalValueConverter : CellValueConverter<decimal>
{
    public override DataCell ConvertToCell(decimal value)
    {
        return new DataCell(value);
    }
}