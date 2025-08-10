using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class TrueBooleanCellValueWriter : BooleanCellValueWriter
{
    private static ReadOnlySpan<byte> EndReferenceTrueBooleanValue => "\" t=\"b\"><v>1</v></c>"u8;
    private static ReadOnlySpan<byte> EndStyleTrueBooleanValue => "\"><v>1</v></c>"u8;
    private static ReadOnlySpan<byte> EndFormulaTrueBooleanValue => "</f><v>1</v></c>"u8;

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite("<c t=\"b\"><v>1</v></c>"u8);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginStyledBooleanCell}{styleId.Id}{EndStyleTrueBooleanValue}");
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        if (styleId is { } style)
        {
            return buffer.TryWrite(
                $"{BeginStyledBooleanCell}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{EndFormulaTrueBooleanValue}");
        }

        return buffer.TryWrite($"{BeginBooleanFormulaCell}{formulaText}{EndFormulaTrueBooleanValue}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceTrueBooleanValue}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceBeginStyled}{styleId.Id}{EndStyleTrueBooleanValue}");
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        if (styleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{state}{EndReferenceBeginStyled}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{EndFormulaTrueBooleanValue}");
        }

        return state.Buffer.TryWrite(
            $"{state}{EndReferenceBeginFormula}" +
            $"{formulaText}" +
            $"{EndFormulaTrueBooleanValue}");
    }

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        return buffer.TryWrite("1"u8);
    }
}
