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
            return StyledCellSpanHelper.TryWriteCell(cell.DataCell, cell.StyleId, Buffer, out bytesNeeded);
        }

        protected override bool FinishWritingCellValue(in StyledCell cell, ref int cellValueIndex)
        {
            return FinishWritingCellValue(cell.DataCell.Value, ref cellValueIndex);
        }

        protected override int GetBytes(in StyledCell cell, bool assertSize)
        {
            return StyledCellSpanHelper.GetBytes(cell.DataCell, cell.StyleId, Buffer.GetNextSpan(), assertSize);
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
