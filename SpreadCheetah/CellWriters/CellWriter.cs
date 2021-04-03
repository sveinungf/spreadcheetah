using SpreadCheetah.Helpers;
using System;

namespace SpreadCheetah.CellWriters
{
    internal sealed class CellWriter : BaseCellWriter<Cell>
    {
        public CellWriter(SpreadsheetBuffer buffer) : base(buffer)
        {
        }

        protected override bool TryWriteCell(in Cell cell, out int bytesNeeded)
        {
            bytesNeeded = 0;
            var remainingBuffer = Buffer.GetRemainingBuffer();

            // Try with an approximate cell value/formula length
            var cellValueLength = cell.DataCell.Value.Length * Utf8Helper.MaxBytePerChar;

            var formulaChars = cell.Formula?.FormulaText.Length ?? 0;

            if (cell.Formula is null)
            {
                // TODO: Same as for styled cells
            }
            else
            {

            }

            throw new Exception();
        }

        protected override bool FinishWritingCellValue(Cell cell, ref int cellValueIndex)
        {
            throw new NotImplementedException();
        }

        protected override int GetBytes(in Cell cell, Span<byte> bytes, bool assertSize)
        {
            throw new NotImplementedException();
        }

        protected override int GetStartElementBytes(Cell cell, Span<byte> bytes)
        {
            throw new NotImplementedException();
        }

        protected override bool TryWriteEndElement(in Cell cell)
        {
            throw new NotImplementedException();
        }
    }
}
