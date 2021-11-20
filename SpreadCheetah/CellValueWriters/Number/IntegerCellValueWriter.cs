using SpreadCheetah.Helpers;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number;

internal sealed class IntegerCellValueWriter : NumberCellValueWriter
{
    protected override int MaxNumberLength => ValueConstants.IntegerValueMaxCharacters;

    protected override int GetValueBytes(in DataCell cell, Span<byte> destination)
    {
        Utf8Formatter.TryFormat(cell.NumberValue.IntValue, destination, out var bytesWritten);
        return bytesWritten;
    }

    public override bool Equals(in CellValue value, in CellValue other) => value.IntValue == other.IntValue;
    public override int GetHashCodeFor(in CellValue value) => value.IntValue.GetHashCode();
}
