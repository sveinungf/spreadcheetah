using SpreadCheetah.Helpers;
using System;

namespace SpreadCheetah.CellValueWriters.Boolean
{
    internal sealed class NullBooleanCellValueWriter : BooleanCellValueWriter
    {
        public override bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer)
        {
            buffer.Index += SpanHelper.GetBytes(DataCellHelper.NullBooleanCell, buffer.GetNextSpan());
            return true;
        }

        protected override ReadOnlySpan<byte> EndFormulaValueBytes() => FormulaCellHelper.EndFormulaEndCell;

        protected override ReadOnlySpan<byte> EndStyleValueBytes() => StyledCellHelper.EndStyleNullValue;
    }
}
