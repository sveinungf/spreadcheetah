using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class TrueBooleanCellValueWriter : BooleanCellValueWriter
{
    protected override ReadOnlySpan<byte> DataCellBytes() => DataCellHelper.TrueBooleanCell;

    protected override ReadOnlySpan<byte> EndFormulaValueBytes() => FormulaCellHelper.EndFormulaTrueBooleanValue;

    protected override ReadOnlySpan<byte> EndStyleValueBytes() => StyledCellHelper.EndStyleTrueBooleanValue;
}
