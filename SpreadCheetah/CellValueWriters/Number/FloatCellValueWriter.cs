using System;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number
{
    internal sealed class FloatCellValueWriter : NumberCellValueWriter
    {
        // https://stackoverflow.com/questions/1701055/what-is-the-maximum-length-in-chars-needed-to-represent-any-double-value
        protected override int MaxNumberLength => 16;

        protected override int GetValueBytes(in DataCell cell, Span<byte> destination)
        {
            Utf8Formatter.TryFormat(cell.NumberValue.FloatValue, destination, out var bytesWritten);
            return bytesWritten;
        }
    }
}
