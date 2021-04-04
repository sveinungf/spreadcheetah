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
            return DataCellHelper.TryWriteCell(cell, Buffer, out bytesNeeded);
        }

        protected override int GetBytes(in DataCell cell, bool assertSize)
        {
            return DataCellHelper.GetBytes(cell, Buffer.GetNextSpan(), assertSize);
        }

        protected override int GetStartElementBytes(in DataCell cell)
        {
            return DataCellHelper.GetStartElementBytes(cell.DataType, Buffer.GetNextSpan());
        }

        protected override bool TryWriteEndElement(in DataCell cell)
        {
            return DataCellHelper.TryWriteEndElement(cell, Buffer);
        }

        protected override bool FinishWritingCellValue(in DataCell cell, ref int cellValueIndex)
        {
            return FinishWritingCellValue(cell.Value, ref cellValueIndex);
        }
    }
}
