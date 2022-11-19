using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class FalseBooleanCellValueWriter : BooleanCellValueWriter
{
    private static ReadOnlySpan<byte> FalseBooleanCell => "<c t=\"b\"><v>0</v></c>"u8;

    protected override bool TryWriteCell(SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        if (!FalseBooleanCell.TryCopyTo(bytes))
            return false;

        buffer.Advance(FalseBooleanCell.Length);
        return true;
    }

    protected override bool TryWriteEndStyleValue(Span<byte> bytes, out int bytesWritten)
    {
        if (StyledCellHelper.EndStyleFalseBooleanValue.TryCopyTo(bytes))
        {
            bytesWritten = StyledCellHelper.EndStyleFalseBooleanValue.Length;
            return true;
        }

        bytesWritten = 0;
        return false;
    }

    protected override bool TryWriteEndFormulaValue(Span<byte> bytes, out int bytesWritten)
    {
        if (FormulaCellHelper.EndFormulaFalseBooleanValue.TryCopyTo(bytes))
        {
            bytesWritten = FormulaCellHelper.EndFormulaFalseBooleanValue.Length;
            return true;
        }

        bytesWritten = 0;
        return false;
    }
}
