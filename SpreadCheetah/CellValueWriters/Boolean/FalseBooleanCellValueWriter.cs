using SpreadCheetah.Helpers;
using System;

namespace SpreadCheetah.CellValueWriters.Boolean
{
    internal sealed class FalseBooleanCellValueWriter : BooleanCellValueWriter
    {
        public override bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer)
        {
            buffer.Index += SpanHelper.GetBytes(DataCellHelper.FalseBooleanCell, buffer.GetNextSpan());
            return true;
        }

        protected override ReadOnlySpan<byte> EndFormulaValueBytes() => FormulaCellHelper.EndFormulaFalseBooleanValue;

        protected override ReadOnlySpan<byte> EndStyleValueBytes() => StyledCellHelper.EndStyleFalseBooleanValue;
    }
}
