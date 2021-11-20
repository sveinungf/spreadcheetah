using SpreadCheetah.Helpers;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number;

internal sealed class DoubleCellValueWriter : NumberCellValueWriter
{
    protected override int MaxNumberLength => ValueConstants.DoubleValueMaxCharacters;

    protected override int GetValueBytes(in DataCell cell, Span<byte> destination)
    {
        Utf8Formatter.TryFormat(cell.NumberValue.DoubleValue, destination, out var bytesWritten);
        return bytesWritten;
    }

    public override bool Equals(in CellValue value, in CellValue other) => value.DoubleValue == other.DoubleValue;
    public override int GetHashCodeFor(in CellValue value) => value.DoubleValue.GetHashCode();
}
