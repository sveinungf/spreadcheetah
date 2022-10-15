using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class TrueBooleanCellValueWriter : BooleanCellValueWriter
{
    protected override ReadOnlySpan<byte> DataCellBytes() => DataCellHelper.TrueBooleanCell;

    protected override ReadOnlySpan<byte> EndFormulaValueBytes() => FormulaCellHelper.EndFormulaTrueBooleanValue;

    protected override ReadOnlySpan<byte> EndStyleValueBytes() => StyledCellHelper.EndStyleTrueBooleanValue;

    protected override bool TryWriteCell(SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        if (!DataCellHelper.TrueBooleanCell.TryCopyTo(bytes))
            return false;

        buffer.Advance(DataCellHelper.TrueBooleanCell.Length);
        return true;
    }
}
