using SpreadCheetah.Helpers;
using System;

namespace SpreadCheetah.CellValueWriters.Boolean
{
    internal sealed class NullBooleanCellValueWriter : BooleanCellValueWriter
    {
        protected override ReadOnlySpan<byte> DataCellBytes() => DataCellHelper.NullBooleanCell;

        protected override ReadOnlySpan<byte> EndFormulaValueBytes() => FormulaCellHelper.EndFormulaEndCell;

        protected override ReadOnlySpan<byte> EndStyleValueBytes() => StyledCellHelper.EndStyleNullValue;
    }
}
