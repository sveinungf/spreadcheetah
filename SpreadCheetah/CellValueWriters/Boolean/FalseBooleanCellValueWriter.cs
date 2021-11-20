using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellValueWriters.Boolean
{
    internal sealed class FalseBooleanCellValueWriter : BooleanCellValueWriter
    {
        protected override ReadOnlySpan<byte> DataCellBytes() => DataCellHelper.FalseBooleanCell;

        protected override ReadOnlySpan<byte> EndFormulaValueBytes() => FormulaCellHelper.EndFormulaFalseBooleanValue;

        protected override ReadOnlySpan<byte> EndStyleValueBytes() => StyledCellHelper.EndStyleFalseBooleanValue;
    }
}
