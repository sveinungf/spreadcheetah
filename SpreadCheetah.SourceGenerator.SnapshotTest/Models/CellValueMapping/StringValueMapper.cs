using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueMapping;

internal class StringValueMapper : ICellValueMapper<string>
{
    public Cell MapToCell(string value)
    {
        return new Cell(value);
    }
}

internal class IntValueMapper : ICellValueMapper<int?>
{
    public Cell MapToCell(int? value)
    {
        return new Cell(value);
    }
}

internal class DecimalValueMapper : ICellValueMapper<decimal>
{
    public Cell MapToCell(decimal value)
    {
        return new Cell(value);
    }
}

