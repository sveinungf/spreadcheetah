using System;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number
{
    internal sealed class IntegerCellValueWriter : NumberCellValueWriter
    {
        protected override int MaxNumberLength => 11; // -2147483648 (int.MinValue)

        protected override int GetValueBytes(in DataCell cell, Span<byte> destination)
        {
            Utf8Formatter.TryFormat(cell.CellValue.IntValue, destination, out var bytesWritten);
            return bytesWritten;
        }
    }
}
