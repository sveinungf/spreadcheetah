using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class TrueBooleanCellValueWriter : BooleanCellValueWriter
{
    private static ReadOnlySpan<byte> TrueBooleanCell => "<c t=\"b\"><v>1</v></c>"u8;
    private static ReadOnlySpan<byte> EndFormulaTrueBooleanValue => "</f><v>1</v></c>"u8;

    protected override bool TryWriteCell(SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        if (!TrueBooleanCell.TryCopyTo(bytes))
            return false;

        buffer.Advance(TrueBooleanCell.Length);
        return true;
    }

    protected override bool TryWriteEndStyleValue(Span<byte> bytes, out int bytesWritten)
    {
        if (StyledCellHelper.EndStyleTrueBooleanValue.TryCopyTo(bytes))
        {
            bytesWritten = StyledCellHelper.EndStyleTrueBooleanValue.Length;
            return true;
        }

        bytesWritten = 0;
        return false;
    }

    protected override bool TryWriteEndFormulaValue(Span<byte> bytes, out int bytesWritten)
    {
        if (EndFormulaTrueBooleanValue.TryCopyTo(bytes))
        {
            bytesWritten = EndFormulaTrueBooleanValue.Length;
            return true;
        }

        bytesWritten = 0;
        return false;
    }
}
