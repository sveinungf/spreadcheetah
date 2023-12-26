using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class FalseBooleanCellValueWriter : BooleanCellValueWriter
{
    private static ReadOnlySpan<byte> FalseBooleanCell => "<c t=\"b\"><v>0</v></c>"u8;
    private static ReadOnlySpan<byte> EndReferenceFalseBooleanValue => "\" t=\"b\"><v>0</v></c>"u8;
    private static ReadOnlySpan<byte> EndStyleFalseBooleanValue => "\"><v>0</v></c>"u8;
    private static ReadOnlySpan<byte> EndFormulaFalseBooleanValue => "</f><v>0</v></c>"u8;

    protected override bool TryWriteCell(SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{FalseBooleanCell}");
    }

    protected override bool TryWriteCell(StyleId styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginStyledBooleanCell}{styleId.Id}{EndStyleFalseBooleanValue}");
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        if (styleId is { } style)
        {
            return buffer.TryWrite(
                $"{BeginStyledBooleanCell}{style}{FormulaCellHelper.EndStyleBeginFormula}" +
                $"{formulaText}" +
                $"{EndFormulaFalseBooleanValue}");
        }

        return buffer.TryWrite($"{BeginBooleanFormulaCell}{formulaText}{EndFormulaFalseBooleanValue}");
    }

    protected override bool TryWriteCellWithReference(CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceFalseBooleanValue}");
    }

    protected override bool TryWriteCellWithReference(StyleId styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceBeginStyled}{styleId.Id}{EndStyleFalseBooleanValue}");
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        if (styleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{state}{EndReferenceBeginStyled}{style.Id}{FormulaCellHelper.EndStyleBeginFormula}" +
                $"{formulaText}" +
                $"{EndFormulaFalseBooleanValue}");
        }

        return state.Buffer.TryWrite(
            $"{state}{EndReferenceBeginFormula}" +
            $"{formulaText}" +
            $"{EndFormulaFalseBooleanValue}");
    }

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        return cell.Formula is null || buffer.TryWrite($"{EndFormulaFalseBooleanValue}");
    }
}
