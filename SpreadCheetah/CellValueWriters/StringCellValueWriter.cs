using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters;

internal sealed class StringCellValueWriter : CellValueWriter
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

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return buffer.TryWriteWithXmlEncode($"{BeginStringCell}{cell.StringValue}{EndStringCell}");
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWriteWithXmlEncode(
            $"{BeginStyledStringCell}{styleId.Id}{EndStyleBeginInlineString}" +
            $"{cell.StringValue}" +
            $"{EndStringCell}");
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        if (styleId is { } style)
        {
            return buffer.TryWrite(
                $"{BeginStyledStringFormulaCell}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{new RawString(formulaText)}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{cachedValue.StringValue}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return buffer.TryWrite(
            $"{BeginStringFormulaCell}" +
            $"{new RawString(formulaText)}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{cachedValue.StringValue}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return state.Buffer.TryWriteWithXmlEncode($"{state}{EndReferenceBeginString}{cell.StringValue}{EndStringCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return state.Buffer.TryWriteWithXmlEncode(
            $"{state}{EndReferenceBeginStyle}{styleId.Id}{EndStyleBeginInlineString}" +
            $"{cell.StringValue}" +
            $"{EndStringCell}");
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        if (styleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{state}{EndReferenceBeginFormulaCellStyle}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{new RawString(formulaText)}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{cachedValue.StringValue}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return state.Buffer.TryWrite(
            $"{state}{EndReferenceBeginFormula}" +
            $"{new RawString(formulaText)}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{cachedValue.StringValue}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{EndStringCell}");
    }

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        return cell.Formula is null
            ? TryWriteEndElement(buffer)
            : buffer.TryWrite($"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return styleId is { } style
            ? buffer.TryWrite($"{BeginStyledStringCell}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}")
            : buffer.TryWrite($"{BeginStringFormulaCell}");
    }

    public override bool WriteFormulaStartElementWithReference(StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return styleId is { } style
            ? state.Buffer.TryWrite($"{state}{EndReferenceBeginStyle}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}")
            : state.Buffer.TryWrite($"{state}{EndReferenceBeginFormula}");
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginStringCell}");
    }

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{BeginStyledStringCell}{styleId.Id}{EndStyleBeginInlineString}");
    }

    public override bool WriteStartElementWithReference(CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceBeginString}");
    }

    public override bool WriteStartElementWithReference(StyleId styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndReferenceBeginStyle}{styleId.Id}{EndStyleBeginInlineString}");
    }

    public override bool CanWriteValuePieceByPiece(in DataCell cell) => true;

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        return buffer.WriteLongString(cell.StringValue, ref valueIndex);
    }
}
