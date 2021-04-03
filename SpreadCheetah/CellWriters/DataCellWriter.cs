using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellWriters
{
    internal sealed class DataCellWriter : BaseCellWriter<DataCell>
    {
        public DataCellWriter(SpreadsheetBuffer buffer) : base(buffer)
        {
        }

        protected override bool TryWriteCell(in DataCell cell, out int bytesNeeded)
        {
            return DataCellSpanHelper.TryWriteCell(cell, Buffer, out bytesNeeded);
        }

        protected override bool FinishWritingCellValue(in DataCell cell, ref int cellValueIndex)
        {
            return FinishWritingCellValue(cell.Value, ref cellValueIndex);
        }

        protected override int GetBytes(in DataCell cell, bool assertSize)
        {
            return DataCellSpanHelper.GetBytes(cell, Buffer.GetNextSpan(), assertSize);
        }

        protected override int GetStartElementBytes(in DataCell cell)
        {
            return DataCellSpanHelper.GetStartElementBytes(cell.DataType, Buffer.GetNextSpan());
        }

        protected override bool TryWriteEndElement(in DataCell cell)
        {
            return DataCellSpanHelper.TryWriteEndElement(cell, Buffer);
        }
    }
}
