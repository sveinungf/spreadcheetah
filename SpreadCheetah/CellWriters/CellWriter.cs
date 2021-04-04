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
            { Formula: not null } => FormulaCellHelper.TryWriteCell(cell.Formula.Value.FormulaText, cell.DataCell, cell.StyleId, Buffer, out bytesNeeded),
            { StyleId: not null } => StyledCellHelper.TryWriteCell(cell.DataCell, cell.StyleId, Buffer, out bytesNeeded),
            _ => DataCellHelper.TryWriteCell(cell.DataCell, Buffer, out bytesNeeded)
        };

        protected override int GetBytes(in Cell cell, bool assertSize) => cell switch
        {
            { Formula: not null } => FormulaCellHelper.GetBytes(cell.Formula.Value.FormulaText, cell.DataCell, cell.StyleId, Buffer.GetNextSpan(), assertSize),
            { StyleId: not null } => StyledCellHelper.GetBytes(cell.DataCell, cell.StyleId, Buffer.GetNextSpan(), assertSize),
            _ => DataCellHelper.GetBytes(cell.DataCell, Buffer.GetNextSpan(), assertSize)
        };

        protected override int GetStartElementBytes(in Cell cell) => cell switch
        {
            { Formula: not null } => FormulaCellHelper.GetStartElementBytes(cell.DataCell.DataType, cell.StyleId, Buffer.GetNextSpan()),
            { StyleId: not null } => StyledCellHelper.GetStartElementBytes(cell.DataCell, cell.StyleId, Buffer.GetNextSpan()),
            _ => DataCellHelper.GetStartElementBytes(cell.DataCell.DataType, Buffer.GetNextSpan())
        };

        protected override bool TryWriteEndElement(in Cell cell)
        {
            if (cell.Formula is null)
                return DataCellHelper.TryWriteEndElement(cell.DataCell, Buffer);

            var cellEnd = string.IsNullOrEmpty(cell.DataCell.Value)
                ? FormulaCellHelper.EndFormulaEndCell
                : FormulaCellHelper.EndCachedValueEndCell;

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
                var separator = FormulaCellHelper.EndFormulaBeginCachedValue;
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
