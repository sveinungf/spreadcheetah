using SpreadCheetah.Helpers;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number;

internal sealed class FloatCellValueWriter : NumberCellValueWriter
{
    protected override int MaxNumberLength => ValueConstants.FloatValueMaxCharacters;

    protected override int GetValueBytes(in DataCell cell, Span<byte> destination)
    {
        Utf8Formatter.TryFormat(cell.NumberValue.FloatValue, destination, out var bytesWritten);
        return bytesWritten;
    }

    public override bool Equals(in CellValue value, in CellValue other) => value.FloatValue == other.FloatValue;
    public override int GetHashCodeFor(in CellValue value) => value.FloatValue.GetHashCode();
}
