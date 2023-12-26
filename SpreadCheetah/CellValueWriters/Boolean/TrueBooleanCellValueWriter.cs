using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal sealed class TrueBooleanCellValueWriter : BooleanCellValueWriter
{
    private static ReadOnlySpan<byte> TrueBooleanCell => "<c t=\"b\"><v>1</v></c>"u8;
    private static ReadOnlySpan<byte> EndReferenceTrueBooleanValue => "\" t=\"b\"><v>1</v></c>"u8;
    private static ReadOnlySpan<byte> EndStyleTrueBooleanValue => "\"><v>1</v></c>"u8;
    private static ReadOnlySpan<byte> EndFormulaTrueBooleanValue => "</f><v>1</v></c>"u8;

    protected override bool TryWriteCell(SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{TrueBooleanCell}");
    }

    protected override bool TryWriteCell(StyleId styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginStyledBooleanCell}{styleId.Id}{EndStyleTrueBooleanValue}");
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        if (styleId is { } style)
        {
            return buffer.TryWrite(
                $"{BeginStyledBooleanCell}{style}{FormulaCellHelper.EndStyleBeginFormula}" +
                $"{formulaText}" +
                $"{EndFormulaTrueBooleanValue}");
        }

        return buffer.TryWrite($"{BeginBooleanFormulaCell}{formulaText}{EndFormulaTrueBooleanValue}");
    }

    protected override bool TryWriteCellWithReference(CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceTrueBooleanValue}");
    }

    protected override bool TryWriteCellWithReference(StyleId styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceBeginStyled}{styleId.Id}{EndStyleTrueBooleanValue}");
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        if (styleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{state}{EndReferenceBeginStyled}{style.Id}{FormulaCellHelper.EndStyleBeginFormula}" +
                $"{formulaText}" +
                $"{EndFormulaTrueBooleanValue}");
        }

        return state.Buffer.TryWrite(
            $"{state}{EndReferenceBeginFormula}" +
            $"{formulaText}" +
            $"{EndFormulaTrueBooleanValue}");
    }

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        return cell.Formula is null || buffer.TryWrite($"{EndFormulaTrueBooleanValue}");
    }
}
