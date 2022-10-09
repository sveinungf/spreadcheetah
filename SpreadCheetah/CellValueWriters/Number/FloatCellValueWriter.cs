using SpreadCheetah.Helpers;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number;

internal sealed class FloatCellValueWriter : NumberCellValueWriter
{
    protected override int MaxNumberLength => ValueConstants.FloatValueMaxCharacters;

    protected override bool TryWriteValue(in DataCell cell, Span<byte> destination, out int bytesWritten)
    {
        return Utf8Formatter.TryFormat(cell.NumberValue.FloatValue, destination, out bytesWritten);
    }

    public override bool Equals(in CellValue value, in CellValue other) => value.FloatValue == other.FloatValue;
    public override int GetHashCodeFor(in CellValue value) => value.FloatValue.GetHashCode();
}
