using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number;

internal sealed class DoubleCellValueWriter : NumberCellValueWriter
{
    protected override bool TryWriteValue(in DataCell cell, Span<byte> destination, out int bytesWritten)
    {
        return Utf8Formatter.TryFormat(cell.NumberValue.DoubleValue, destination, out bytesWritten);
    }
}
