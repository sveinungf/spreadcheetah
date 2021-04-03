using SpreadCheetah.Helpers;

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

        protected override int GetBytes(in Cell cell, bool assertSize) => cell switch
        {
            { Formula: not null } => FormulaCellSpanHelper.GetBytes(cell.Formula.Value.FormulaText, cell.DataCell, cell.StyleId, Buffer.GetNextSpan(), assertSize),
            { StyleId: not null } => StyledCellSpanHelper.GetBytes(cell.DataCell, cell.StyleId, Buffer.GetNextSpan(), assertSize),
            _ => DataCellSpanHelper.GetBytes(cell.DataCell, Buffer.GetNextSpan(), assertSize)
        };

        protected override int GetStartElementBytes(in Cell cell) => cell switch
        {
            { Formula: not null } => FormulaCellSpanHelper.GetStartElementBytes(cell.DataCell.DataType, cell.StyleId, Buffer.GetNextSpan()),
            { StyleId: not null } => StyledCellSpanHelper.GetStartElementBytes(cell.DataCell, cell.StyleId, Buffer.GetNextSpan()),
            _ => DataCellSpanHelper.GetStartElementBytes(cell.DataCell.DataType, Buffer.GetNextSpan())
        };

        protected override bool TryWriteEndElement(in Cell cell)
        {
            if (cell.Formula is null)
                return DataCellSpanHelper.TryWriteEndElement(cell.DataCell, Buffer);

            var cellEnd = string.IsNullOrEmpty(cell.DataCell.Value)
                ? FormulaCellSpanHelper.EndFormulaEndCell
                : FormulaCellSpanHelper.EndCachedValueEndCell;

            if (cellEnd.Length > Buffer.GetRemainingBuffer())
                return false;

            Buffer.Index += SpanHelper.GetBytes(cellEnd, Buffer.GetNextSpan());
            return true;
        }

        protected override bool FinishWritingCellValue(in Cell cell, ref int cellValueIndex)
        {
            if (cell.Formula is null)
                return FinishWritingCellValue(cell.DataCell.Value, ref cellValueIndex);

            var formulaText = cell.Formula.Value.FormulaText;

            // Write the formula
            if (cellValueIndex < formulaText.Length)
            {
                // If there is a cached value, we need to write "[FORMULA]</f><v>[CACHEDVALUE]"
                if (!FinishWritingCellValue(formulaText, ref cellValueIndex)) return false;

                // Otherwise, we only need to write the formula
                if (string.IsNullOrEmpty(cell.DataCell.Value)) return true;
            }

            var cachedValueStartIndex = formulaText.Length + 1;

            // Write the "</f><v>" part
            if (cellValueIndex < cachedValueStartIndex)
            {
                var separator = FormulaCellSpanHelper.EndFormulaBeginCachedValue;
                if (separator.Length > Buffer.GetRemainingBuffer()) return false;
                Buffer.Index += SpanHelper.GetBytes(separator, Buffer.GetNextSpan());
                cellValueIndex = cachedValueStartIndex;
            }

            // Write the cached value
            var cachedValueIndex = cellValueIndex - cachedValueStartIndex;
            var result = FinishWritingCellValue(cell.DataCell.Value, ref cachedValueIndex);
            cellValueIndex = cachedValueIndex + cachedValueStartIndex;
            return result;
        }
    }
}
