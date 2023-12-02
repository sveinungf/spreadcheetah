using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number;

internal sealed class FloatCellValueWriter : NumberCellValueWriter
{
    protected override bool TryWriteValue(in DataCell cell, Span<byte> destination, out int bytesWritten)
    {
        return Utf8Formatter.TryFormat(cell.NumberValue.FloatValue, destination, out bytesWritten);
    }
}
