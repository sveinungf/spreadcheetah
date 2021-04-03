using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellWriters
{
    internal sealed class StyledCellWriter : BaseCellWriter<StyledCell>
    {
        public StyledCellWriter(SpreadsheetBuffer buffer) : base(buffer)
        {
        }

        protected override bool TryWriteCell(in StyledCell cell, out int bytesNeeded)
        {
            bytesNeeded = 0;
            var remainingBuffer = Buffer.GetRemainingBuffer();

            // Try with an approximate cell value length
            var cellValueLength = cell.DataCell.Value.Length * Utf8Helper.MaxBytePerChar;
            if (StyledCellSpanHelper.MaxCellElementLength + cellValueLength < remainingBuffer)
            {
                Buffer.Index += StyledCellSpanHelper.GetBytes(cell, Buffer.GetNextSpan(), false);
                return true;
            }

            // Try with a more accurate cell value length
            cellValueLength = Utf8Helper.GetByteCount(cell.DataCell.Value);
            bytesNeeded = StyledCellSpanHelper.MaxCellElementLength + cellValueLength;
            if (bytesNeeded < remainingBuffer)
            {
                Buffer.Index += StyledCellSpanHelper.GetBytes(cell, Buffer.GetNextSpan(), false);
                return true;
            }

            return false;
        }

        protected override bool FinishWritingCellValue(in StyledCell cell, ref int cellValueIndex)
        {
            return FinishWritingCellValue(cell.DataCell.Value, ref cellValueIndex);
        }

        protected override int GetBytes(in StyledCell cell, bool assertSize)
        {
            return StyledCellSpanHelper.GetBytes(cell, Buffer.GetNextSpan(), assertSize);
        }

        protected override int GetStartElementBytes(in StyledCell cell)
        {
            return StyledCellSpanHelper.GetStartElementBytes(cell, Buffer.GetNextSpan());
        }

        protected override bool TryWriteEndElement(in StyledCell cell)
        {
            return DataCellSpanHelper.TryWriteEndElement(cell.DataCell, Buffer);
        }
    }
}
