using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class FalseBooleanCellValueWriter : BooleanCellValueWriter
{
    private static ReadOnlySpan<byte> EndReferenceFalseBooleanValue => "\" t=\"b\"><v>0</v></c>"u8;
    private static ReadOnlySpan<byte> EndStyleFalseBooleanValue => "\"><v>0</v></c>"u8;
    private static ReadOnlySpan<byte> EndFormulaFalseBooleanValue => "</f><v>0</v></c>"u8;

    public override bool TryWriteCell(in DataCell cell, CellWriterState state)
    {
        return state.Buffer.TryWrite("<c t=\"b\"><v>0</v></c>"u8);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginStyledBooleanCell}{styleId.Id}{EndStyleFalseBooleanValue}");
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, CellWriterState state)
    {
        if (styleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{BeginStyledBooleanCell}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{EndFormulaFalseBooleanValue}");
        }

        return state.Buffer.TryWrite($"{BeginBooleanFormulaCell}{formulaText}{EndFormulaFalseBooleanValue}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceFalseBooleanValue}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceBeginStyled}{styleId.Id}{EndStyleFalseBooleanValue}");
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, CellWriterState state)
    {
        if (styleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{state}{EndReferenceBeginStyled}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{EndFormulaFalseBooleanValue}");
        }

        return state.Buffer.TryWrite(
            $"{state}{EndReferenceBeginFormula}" +
            $"{formulaText}" +
            $"{EndFormulaFalseBooleanValue}");
    }

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        return buffer.TryWrite("0"u8);
    }
}
