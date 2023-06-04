using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class TrueBooleanCellValueWriter : BooleanCellValueWriter
{
    private static ReadOnlySpan<byte> TrueBooleanCell => "<c t=\"b\"><v>1</v></c>"u8;
    private static ReadOnlySpan<byte> EndStyleTrueBooleanValue => "\"><v>1</v></c>"u8;
    private static ReadOnlySpan<byte> EndFormulaTrueBooleanValue => "</f><v>1</v></c>"u8;

    protected override bool TryWriteCell(CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!state.WriteCellReferenceAttributes)
        {
            if (!TrueBooleanCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!"<c r=\""u8.TryCopyTo(bytes, ref written)) return false;
            if (!SpanHelper.TryWriteCellReference(state.Column + 1, state.NextRowIndex - 1, bytes, ref written)) return false;
            if (!"\" t=\"b\"><v>1</v></c>"u8.TryCopyTo(bytes, ref written)) return false;
        }

        buffer.Advance(written);
        return true;
    }

    protected override bool TryWriteEndStyleValue(Span<byte> bytes, out int bytesWritten)
    {
        if (EndStyleTrueBooleanValue.TryCopyTo(bytes))
        {
            bytesWritten = EndStyleTrueBooleanValue.Length;
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
