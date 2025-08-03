using SpreadCheetah.CellValues;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Characters;

internal abstract class StringCellValueWriterBase : CellValueWriter
{
    private static ReadOnlySpan<byte> BeginStringCell => "<c t=\"inlineStr\"><is><t>"u8;
    private static ReadOnlySpan<byte> BeginStyledStringCell => "<c t=\"inlineStr\" s=\""u8;
    private static ReadOnlySpan<byte> BeginStringFormulaCell => "<c t=\"str\"><f>"u8;
    private static ReadOnlySpan<byte> BeginStyledStringFormulaCell => "<c t=\"str\" s=\""u8;
    private static ReadOnlySpan<byte> EndReferenceBeginString => "\" t=\"inlineStr\"><is><t>"u8;
    private static ReadOnlySpan<byte> EndReferenceBeginStyle => "\" t=\"inlineStr\" s=\""u8;
    private static ReadOnlySpan<byte> EndReferenceBeginFormula => "\" t=\"str\"><f>"u8;
    private static ReadOnlySpan<byte> EndReferenceBeginFormulaCellStyle => "\" t=\"str\" s=\""u8;
    private static ReadOnlySpan<byte> EndStyleBeginInlineString => "\"><is><t>"u8;
    private static ReadOnlySpan<byte> EndStringCell => "</t></is></c>"u8;

    protected abstract ReadOnlySpan<char> GetSpan(in CellValue cell);

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginStringCell}{GetSpan(cell.Value)}{EndStringCell}");
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite(
            $"{BeginStyledStringCell}{styleId.Id}{EndStyleBeginInlineString}" +
            $"{GetSpan(cell.Value)}" +
            $"{EndStringCell}");
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        if (styleId is { } style)
        {
            return buffer.TryWrite(
                $"{BeginStyledStringFormulaCell}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{GetSpan(cachedValue.Value)}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return buffer.TryWrite(
            $"{BeginStringFormulaCell}" +
            $"{formulaText}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{GetSpan(cachedValue.Value)}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceBeginString}{GetSpan(cell.Value)}{EndStringCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite(
            $"{state}{EndReferenceBeginStyle}{styleId.Id}{EndStyleBeginInlineString}" +
            $"{GetSpan(cell.Value)}" +
            $"{EndStringCell}");
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        if (styleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{state}{EndReferenceBeginFormulaCellStyle}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{GetSpan(cachedValue.Value)}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return state.Buffer.TryWrite(
            $"{state}{EndReferenceBeginFormula}" +
            $"{formulaText}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{GetSpan(cachedValue.Value)}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite(EndStringCell);
    }

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        return cell.Formula is null
            ? TryWriteEndElement(buffer)
            : buffer.TryWrite(FormulaCellHelper.EndCachedValueEndCell);
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return styleId is { } style
            ? buffer.TryWrite($"{BeginStyledStringCell}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}")
            : buffer.TryWrite(BeginStringFormulaCell);
    }

    public override bool WriteFormulaStartElementWithReference(StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return styleId is { } style
            ? state.Buffer.TryWrite($"{state}{EndReferenceBeginStyle}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}")
            : state.Buffer.TryWrite($"{state}{EndReferenceBeginFormula}");
    }

    public override bool WriteStartElement(StyleId? styleId, SpreadsheetBuffer buffer)
    {
        return styleId is null
            ? buffer.TryWrite(BeginStringCell)
            : buffer.TryWrite($"{BeginStyledStringCell}{styleId.Id}{EndStyleBeginInlineString}");
    }

    public override bool WriteStartElementWithReference(StyleId? styleId, CellWriterState state)
    {
        return styleId is null
            ? state.Buffer.TryWrite($"{state}{EndReferenceBeginString}")
            : state.Buffer.TryWrite($"{state}{EndReferenceBeginStyle}{styleId.Id}{EndStyleBeginInlineString}");
    }

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        return buffer.WriteLongString(GetSpan(cell.Value), ref valueIndex);
    }
}
