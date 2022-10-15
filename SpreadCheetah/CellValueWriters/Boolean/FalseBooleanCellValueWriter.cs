using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class FalseBooleanCellValueWriter : BooleanCellValueWriter
{
    protected override ReadOnlySpan<byte> EndFormulaValueBytes() => FormulaCellHelper.EndFormulaFalseBooleanValue;

    protected override ReadOnlySpan<byte> EndStyleValueBytes() => StyledCellHelper.EndStyleFalseBooleanValue;

    protected override bool TryWriteCell(SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        if (!DataCellHelper.FalseBooleanCell.TryCopyTo(bytes))
            return false;

        buffer.Advance(DataCellHelper.FalseBooleanCell.Length);
        return true;
    }
}
