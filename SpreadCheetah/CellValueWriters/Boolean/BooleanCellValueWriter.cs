using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal abstract class BooleanCellValueWriter : CellValueWriter
{
    private static ReadOnlySpan<byte> BeginStyledBooleanCell => "<c t=\"b\" s=\""u8;
    private static ReadOnlySpan<byte> BeginBooleanFormulaCell => "<c t=\"b\"><f>"u8;

    [Obsolete]
    protected abstract bool TryWriteCell(CellWriterState state);
    protected abstract bool TryWriteCell(SpreadsheetBuffer buffer);
    protected abstract bool TryWriteCellWithReference(CellWriterState state);
    protected abstract bool TryWriteEndStyleValue(Span<byte> bytes, out int bytesWritten);
    protected abstract bool TryWriteEndFormulaValue(Span<byte> bytes, out int bytesWritten);

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return TryWriteCell(state);
    }

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(buffer);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return TryWriteCell(styleId, state);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(styleId, buffer);
    }

    private bool TryWriteCell(StyleId styleId, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginStyledBooleanCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"b\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!TryWriteEndStyleValue(bytes.Slice(written), out var endLength)) return false;
        written += endLength;

        buffer.Advance(written);
        return true;
    }

    private bool TryWriteCell(StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!BeginStyledBooleanCell.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!TryWriteEndStyleValue(bytes.Slice(written), out var endLength)) return false;
        written += endLength;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();

        if (!TryWriteFormulaCellStart(styleId, state, bytes, out var written)) return false;
        if (!SpanHelper.TryWrite(formulaText, bytes, ref written)) return false;
        if (!TryWriteEndFormulaValue(bytes.Slice(written), out var endLength)) return false;

        buffer.Advance(written + endLength);
        return true;
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var written = 0;

        if (styleId is null)
        {
            if (!BeginBooleanFormulaCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!BeginStyledBooleanCell.TryCopyTo(bytes, ref written)) return false;
            if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
            if (!FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(formulaText, bytes, ref written)) return false;
        if (!TryWriteEndFormulaValue(bytes.Slice(written), out var endLength)) return false;

        buffer.Advance(written + endLength);
        return true;
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return TryWriteCellWithReference(state);
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return TryWriteCellWithReference(styleId, state);
    }

    private bool TryWriteCellWithReference(StyleId styleId, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
        if (!"\" t=\"b\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!TryWriteEndStyleValue(bytes.Slice(written), out var endLength)) return false;
        written += endLength;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (styleId is null)
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"b\"><f>"u8.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"b\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
            if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
            if (!FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(formulaText, bytes, ref written)) return false;
        if (!TryWriteEndFormulaValue(bytes.Slice(written), out var endLength)) return false;

        buffer.Advance(written + endLength);
        return true;
    }

    private static bool TryWriteFormulaCellStart(StyleId? styleId, CellWriterState state, Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;
        var written = 0;

        if (styleId is null)
        {
            if (!state.WriteCellReferenceAttributes)
                return BeginBooleanFormulaCell.TryCopyTo(bytes, ref bytesWritten);

            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"b\"><f>"u8.TryCopyTo(bytes, ref written)) return false;

            bytesWritten = written;
            return true;
        }

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginStyledBooleanCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"b\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes, ref written)) return false;

        bytesWritten = written;
        return true;
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer) => true;

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        if (cell.Formula is null)
            return true;

        var bytes = buffer.GetSpan();
        if (TryWriteEndFormulaValue(bytes, out var bytesWritten))
        {
            buffer.Advance(bytesWritten);
            return true;
        }

        return false;
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (styleId is null)
        {
            if (!state.WriteCellReferenceAttributes)
            {
                if (!BeginBooleanFormulaCell.TryCopyTo(bytes, ref written)) return false;
                buffer.Advance(written);
                return true;
            }

            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"b\"><f>"u8.TryCopyTo(bytes, ref written)) return false;

            buffer.Advance(written);
            return true;
        }

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginStyledBooleanCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"b\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool WriteStartElement(CellWriterState state) => TryWriteCell(state);

    public override bool WriteStartElement(StyleId styleId, CellWriterState state) => TryWriteCell(styleId, state);
    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer) => TryWriteCell(styleId, buffer);
    public override bool WriteStartElementWithReference(StyleId styleId, CellWriterState state) => TryWriteCellWithReference(styleId, state);

    /// <summary>
    /// Returns false because the value is written together with the end element in <see cref="TryWriteEndElement(in Cell, SpreadsheetBuffer)"/>.
    /// </summary>
    public override bool CanWriteValuePieceByPiece(in DataCell cell) => false;
    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex) => true;
}
