namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class FalseBooleanCellValueWriter : BooleanCellValueWriter
{
    private static ReadOnlySpan<byte> FalseBooleanCell => "<c t=\"b\"><v>0</v></c>"u8;
    private static ReadOnlySpan<byte> EndStyleFalseBooleanValue => "\"><v>0</v></c>"u8;
    private static ReadOnlySpan<byte> EndFormulaFalseBooleanValue => "</f><v>0</v></c>"u8;

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
        if (EndStyleFalseBooleanValue.TryCopyTo(bytes))
        {
            bytesWritten = EndStyleFalseBooleanValue.Length;
            return true;
        }

        bytesWritten = 0;
        return false;
    }

    protected override bool TryWriteEndFormulaValue(Span<byte> bytes, out int bytesWritten)
    {
        if (EndFormulaFalseBooleanValue.TryCopyTo(bytes))
        {
            bytesWritten = EndFormulaFalseBooleanValue.Length;
            return true;
        }

        bytesWritten = 0;
        return false;
    }
}
