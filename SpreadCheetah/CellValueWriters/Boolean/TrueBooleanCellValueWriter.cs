using SpreadCheetah.Helpers;
using System;

namespace SpreadCheetah.CellValueWriters.Boolean
{
    internal sealed class TrueBooleanCellValueWriter : BooleanCellValueWriter
    {
        public override bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer)
        {
            buffer.Index += SpanHelper.GetBytes(DataCellHelper.TrueBooleanCell, buffer.GetNextSpan());
            return true;
        }

        protected override ReadOnlySpan<byte> EndFormulaValueBytes() => FormulaCellHelper.EndFormulaTrueBooleanValue;

        protected override ReadOnlySpan<byte> EndStyleValueBytes() => StyledCellHelper.EndStyleTrueBooleanValue;
    }
}
