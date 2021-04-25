namespace SpreadCheetah.CellWriters
{
    internal sealed class DataCellWriter : BaseCellWriter<DataCell>
    {
        public DataCellWriter(SpreadsheetBuffer buffer) : base(buffer)
        {
        }

        protected override bool TryWriteCell(in DataCell cell, out int bytesNeeded)
        {
            return cell.Writer.TryWriteCell(cell, Buffer, out bytesNeeded);
        }

        protected override bool GetBytes(in DataCell cell, bool assertSize)
        {
            return cell.Writer.GetBytes(cell, Buffer);
        }

        protected override bool WriteStartElement(in DataCell cell)
        {
            return cell.Writer.WriteStartElement(Buffer);
        }

        protected override bool TryWriteEndElement(in DataCell cell)
        {
            return cell.Writer.TryWriteEndElement(Buffer);
        }

        protected override bool FinishWritingCellValue(in DataCell cell, ref int cellValueIndex)
        {
            return FinishWritingCellValue(cell.StringValue!, ref cellValueIndex);
        }
    }
}
