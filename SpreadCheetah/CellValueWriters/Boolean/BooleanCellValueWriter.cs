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

    protected abstract bool TryWriteCell(SpreadsheetBuffer buffer);
    protected abstract bool TryWriteCell(StyleId styleId, SpreadsheetBuffer buffer);
    protected abstract bool TryWriteCellWithReference(CellWriterState state);
    protected abstract bool TryWriteCellWithReference(StyleId styleId, CellWriterState state);

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(buffer);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(styleId, buffer);
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return TryWriteCellWithReference(state);
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return TryWriteCellWithReference(styleId, state);
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer) => true;

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return styleId is { } style
            ? buffer.TryWrite($"{BeginStyledBooleanCell}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}")
            : buffer.TryWrite($"{BeginBooleanFormulaCell}");
    }

    public override bool WriteFormulaStartElementWithReference(StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return styleId is { } style
            ? state.Buffer.TryWrite($"{state}{EndReferenceBeginStyled}{style.Id}{FormulaCellHelper.EndQuoteBeginFormula}")
            : state.Buffer.TryWrite($"{state}{EndReferenceBeginFormula}");
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer) => TryWriteCell(buffer);
    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer) => TryWriteCell(styleId, buffer);
    public override bool WriteStartElementWithReference(CellWriterState state) => TryWriteCellWithReference(state);
    public override bool WriteStartElementWithReference(StyleId styleId, CellWriterState state) => TryWriteCellWithReference(styleId, state);

    /// <summary>
    /// Returns false because the value is written together with the end element.
    /// </summary>
    public override bool CanWriteValuePieceByPiece(in DataCell cell) => false;
    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex) => true;
}
