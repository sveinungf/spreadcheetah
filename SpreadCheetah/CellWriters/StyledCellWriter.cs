namespace SpreadCheetah.CellWriters
{
    internal sealed class StyledCellWriter : BaseCellWriter<StyledCell>
    {
        public StyledCellWriter(SpreadsheetBuffer buffer) : base(buffer)
        {
        }

        protected override bool TryWriteCell(in StyledCell cell, out int bytesNeeded)
        {
            return cell.StyleId is null
                ? cell.DataCell.Writer.TryWriteCell(cell.DataCell, Buffer, out bytesNeeded)
                : cell.DataCell.Writer.TryWriteCell(cell.DataCell, cell.StyleId, Buffer, out bytesNeeded);
        }

        protected override bool GetBytes(in StyledCell cell, bool assertSize)
        {
            return cell.StyleId is null
                ? cell.DataCell.Writer.GetBytes(cell.DataCell, Buffer)
                : cell.DataCell.Writer.GetBytes(cell.DataCell, cell.StyleId, Buffer);
        }

        protected override bool WriteStartElement(in StyledCell cell)
        {
            return cell.StyleId is null
                ? cell.DataCell.Writer.WriteStartElement(cell.DataCell, Buffer)
                : cell.DataCell.Writer.WriteStartElement(cell.DataCell, cell.StyleId, Buffer);
        }

        protected override bool TryWriteEndElement(in StyledCell cell)
        {
            return cell.DataCell.Writer.TryWriteEndElement(cell.DataCell, Buffer);
        }

        protected override bool FinishWritingCellValue(in StyledCell cell, ref int cellValueIndex)
        {
            return FinishWritingCellValue(cell.DataCell.StringValue!, ref cellValueIndex);
        }
    }
}
