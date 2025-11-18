using SpreadCheetah.CellValues;

namespace SpreadCheetah.CellValueWriters.Characters;

internal sealed class StringCellValueWriter : StringCellValueWriterBase
{
    protected override ReadOnlySpan<char> GetSpan(in CellValue cell) => cell.StringOrPrimitive.StringValue;
}
