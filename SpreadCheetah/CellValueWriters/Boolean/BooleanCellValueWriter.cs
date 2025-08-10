using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal abstract class BooleanCellValueWriter : CellValueWriter
{
    protected static ReadOnlySpan<byte> BeginStyledBooleanCell => "<c t=\"b\" s=\""u8;
    protected static ReadOnlySpan<byte> BeginBooleanFormulaCell => "<c t=\"b\"><f>"u8;
    protected static ReadOnlySpan<byte> EndReferenceBeginStyled => "\" t=\"b\" s=\""u8;
    protected static ReadOnlySpan<byte> EndReferenceBeginFormula => "\" t=\"b\"><f>"u8;

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer) => buffer.TryWrite("</v></c>"u8);

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer) => TryWriteEndElement(buffer);

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return styleId is { } style
            ? buffer.TryWrite($"{BeginStyledBooleanCell}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}")
            : buffer.TryWrite(BeginBooleanFormulaCell);
    }

    public override bool WriteFormulaStartElementWithReference(StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return styleId is { } style
            ? state.Buffer.TryWrite($"{state}{EndReferenceBeginStyled}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}")
            : state.Buffer.TryWrite($"{state}{EndReferenceBeginFormula}");
    }
}
