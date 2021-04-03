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
            throw new NotImplementedException();
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
