using SpreadCheetah.Helpers;
using System;

namespace SpreadCheetah.CellWriters
{
    internal sealed class CellWriter : BaseCellWriter<Cell>
    {
        public CellWriter(SpreadsheetBuffer buffer) : base(buffer)
        {
        }

        protected override bool TryWriteCell(in Cell cell, out int bytesNeeded) => cell switch
        {
            { Formula: not null } => FormulaCellSpanHelper.TryWriteCell(cell.Formula.Value.FormulaText, cell.DataCell, cell.StyleId, Buffer, out bytesNeeded),
            { StyleId: not null } => StyledCellSpanHelper.TryWriteCell(cell.DataCell, cell.StyleId, Buffer, out bytesNeeded),
            _ => DataCellSpanHelper.TryWriteCell(cell.DataCell, Buffer, out bytesNeeded)
        };

        protected override bool FinishWritingCellValue(in Cell cell, ref int cellValueIndex)
        {
            if (cell.Formula is null)
                return FinishWritingCellValue(cell.DataCell.Value, ref cellValueIndex);

            var formulaText = cell.Formula.Value.FormulaText;

            // Write the formula
            if (cellValueIndex < formulaText.Length)
            {
                var formulaWritten = FinishWritingCellValue(formulaText, ref cellValueIndex);

                // If no cached value, we only need to write the formula
                if (string.IsNullOrEmpty(cell.DataCell.Value)) return formulaWritten;

                // Otherwise, we need to also write "</f><v>[CACHEDVALUE]"
                if (!formulaWritten) return false;
            }

            var separator = FormulaCellSpanHelper.EndFormulaBeginCachedValue;
            var cachedValueStartIndex = formulaText.Length + separator.Length;

            // Write the "</f><v>" part
            if (cellValueIndex < cachedValueStartIndex)
            {
                if (separator.Length > Buffer.GetRemainingBuffer()) return false;
                Buffer.Index += SpanHelper.GetBytes(separator, Buffer.GetNextSpan());
                cellValueIndex += separator.Length;
            }

            // Write the cached value
            var cachedValueIndex = cellValueIndex - cachedValueStartIndex;
            var result = FinishWritingCellValue(cell.DataCell.Value, ref cachedValueIndex);
            cellValueIndex = cachedValueStartIndex + cachedValueIndex;
            return result;
        }

        protected override int GetBytes(in Cell cell, bool assertSize)
        {
            throw new NotImplementedException();
        }

        protected override int GetStartElementBytes(in Cell cell)
        {
            throw new NotImplementedException();
        }

        protected override bool TryWriteEndElement(in Cell cell)
        {
            throw new NotImplementedException();
        }
    }
}
