using SpreadCheetah.Helpers;
using System;

namespace SpreadCheetah.CellWriters
{
    internal sealed class DataCellWriter : BaseCellWriter<Cell>
    {
        public DataCellWriter(SpreadsheetBuffer buffer) : base(buffer)
        {
        }

        protected override bool TryWriteCell(in Cell cell, out int bytesNeeded)
        {
            bytesNeeded = 0;
            var remainingBuffer = Buffer.GetRemainingBuffer();

            // Try with an approximate cell value length
            var cellValueLength = cell.Value.Length * Utf8Helper.MaxBytePerChar;
            if (CellSpanHelper.MaxCellElementLength + cellValueLength < remainingBuffer)
            {
                Buffer.Index += CellSpanHelper.GetBytes(cell, Buffer.GetNextSpan(), false);
                return true;
            }

            // Try with a more accurate cell value length
            cellValueLength = Utf8Helper.GetByteCount(cell.Value);
            bytesNeeded = CellSpanHelper.MaxCellElementLength + cellValueLength;
            if (bytesNeeded < remainingBuffer)
            {
                Buffer.Index += CellSpanHelper.GetBytes(cell, Buffer.GetNextSpan(), false);
                return true;
            }

            return false;
        }

        protected override bool FinishWritingCellValue(Cell cell, ref int cellValueIndex)
        {
            return FinishWritingCellValue(cell.Value, ref cellValueIndex);
        }

        protected override int GetBytes(in Cell cell, Span<byte> bytes, bool assertSize)
        {
            return CellSpanHelper.GetBytes(cell, Buffer.GetNextSpan(), assertSize);
        }

        protected override int GetStartElementBytes(Cell cell, Span<byte> bytes)
        {
            return CellSpanHelper.GetStartElementBytes(cell.DataType, Buffer.GetNextSpan());
        }

        protected override int GetEndElementBytes(Cell cell, Span<byte> bytes)
        {
            return CellSpanHelper.GetEndElementBytes(cell.DataType, Buffer.GetNextSpan());
        }
    }
}
