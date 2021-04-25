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
            { Formula: not null } => cell.DataCell.Writer.TryWriteCell(cell.Formula.Value.FormulaText, cell.DataCell, cell.StyleId, Buffer, out bytesNeeded),
            { StyleId: not null } => cell.DataCell.Writer.TryWriteCell(cell.DataCell, cell.StyleId, Buffer, out bytesNeeded),
            _ => cell.DataCell.Writer.TryWriteCell(cell.DataCell, Buffer, out bytesNeeded)
        };

        protected override bool GetBytes(in Cell cell, bool assertSize) => cell switch
        {
            { Formula: not null } => cell.DataCell.Writer.GetBytes(cell.Formula.Value.FormulaText, cell.DataCell, cell.StyleId, Buffer),
            { StyleId: not null } => cell.DataCell.Writer.GetBytes(cell.DataCell, cell.StyleId, Buffer),
            _ => cell.DataCell.Writer.GetBytes(cell.DataCell, Buffer)
        };

        protected override bool WriteStartElement(in Cell cell) => cell switch
        {
            { Formula: not null } => cell.DataCell.Writer.WriteFormulaStartElement(cell.StyleId, Buffer),
            { StyleId: not null } => cell.DataCell.Writer.WriteStartElement(cell.StyleId, Buffer),
            _ => cell.DataCell.Writer.WriteStartElement(Buffer)
        };

        protected override bool TryWriteEndElement(in Cell cell)
        {
            return cell.DataCell.Writer.TryWriteEndElement(cell, Buffer);
        }

        protected override bool FinishWritingCellValue(in Cell cell, ref int cellValueIndex)
        {
            if (cell.Formula is null)
                return FinishWritingCellValue(cell.DataCell.StringValue!, ref cellValueIndex);

            var formulaText = cell.Formula.Value.FormulaText;

            // Write the formula
            if (cellValueIndex < formulaText.Length)
            {
                // If there is a cached value, we need to write "[FORMULA]</f><v>[CACHEDVALUE]"
                if (!FinishWritingCellValue(formulaText, ref cellValueIndex)) return false;

                // Otherwise, we only need to write the formula
                if (string.IsNullOrEmpty(cell.DataCell.StringValue)) return true;
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
            var result = FinishWritingCellValue(cell.DataCell.StringValue!, ref cachedValueIndex);
            cellValueIndex = cachedValueIndex + cachedValueStartIndex;
            return result;
        }
    }
}
