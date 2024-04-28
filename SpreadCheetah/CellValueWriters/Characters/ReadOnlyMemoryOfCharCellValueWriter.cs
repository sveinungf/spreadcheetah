using SpreadCheetah.CellValues;

namespace SpreadCheetah.CellValueWriters.Characters;

internal sealed class ReadOnlyMemoryOfCharCellValueWriter : StringCellValueWriterBase
{
    protected override ReadOnlySpan<char> GetSpan(in CellValue cell) => cell.Memory.Span;
}
