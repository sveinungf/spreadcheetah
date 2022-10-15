using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class TrueBooleanCellValueWriter : BooleanCellValueWriter
{
    protected override bool TryWriteCell(SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        if (!DataCellHelper.TrueBooleanCell.TryCopyTo(bytes))
            return false;

        buffer.Advance(DataCellHelper.TrueBooleanCell.Length);
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
        if (FormulaCellHelper.EndFormulaTrueBooleanValue.TryCopyTo(bytes))
        {
            bytesWritten = FormulaCellHelper.EndFormulaTrueBooleanValue.Length;
            return true;
        }

        bytesWritten = 0;
        return false;
    }
}
