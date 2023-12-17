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
    private static ReadOnlySpan<byte> EndStyleBeginInlineString => "\"><is><t>"u8;
    private static ReadOnlySpan<byte> EndStringCell => "</t></is></c>"u8;

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginStringCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"inlineStr\"><is><t>"u8.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(cell.StringValue, bytes, ref written)) return false;
        if (!EndStringCell.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!BeginStringCell.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(cell.StringValue, bytes, ref written)) return false;
        if (!EndStringCell.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginStyledStringCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"inlineStr\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!EndStyleBeginInlineString.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(cell.StringValue, bytes, ref written)) return false;
        if (!EndStringCell.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!BeginStyledStringCell.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!EndStyleBeginInlineString.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(cell.StringValue, bytes, ref written)) return false;
        if (!EndStringCell.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();

        if (!TryWriteFormulaCellStart(styleId, state, bytes, out var written)) return false;
        if (!SpanHelper.TryWrite(formulaText, bytes, ref written)) return false;
        if (!FormulaCellHelper.EndFormulaBeginCachedValue.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(cachedValue.StringValue, bytes, ref written)) return false;
        if (!FormulaCellHelper.EndCachedValueEndCell.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
        if (!"\" t=\"inlineStr\"><is><t>"u8.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(cell.StringValue, bytes, ref written)) return false;
        if (!EndStringCell.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
        if (!"\" t=\"inlineStr\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!EndStyleBeginInlineString.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(cell.StringValue, bytes, ref written)) return false;
        if (!EndStringCell.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    private static bool TryWriteFormulaCellStart(StyleId? styleId, CellWriterState state, Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;
        var written = 0;

        if (styleId is null)
        {
            if (!state.WriteCellReferenceAttributes)
                return BeginStringFormulaCell.TryCopyTo(bytes, ref bytesWritten);

            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"str\"><f>"u8.TryCopyTo(bytes, ref written)) return false;

            bytesWritten = written;
            return true;
        }

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginStyledStringFormulaCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"str\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes, ref written)) return false;

        bytesWritten = written;
        return true;
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        if (!EndStringCell.TryCopyTo(bytes))
            return false;

        buffer.Advance(EndStringCell.Length);
        return true;
    }

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        if (cell.Formula is null)
            return TryWriteEndElement(buffer);

        var bytes = buffer.GetSpan();
        if (!FormulaCellHelper.EndCachedValueEndCell.TryCopyTo(bytes))
            return false;

        buffer.Advance(FormulaCellHelper.EndCachedValueEndCell.Length);
        return true;
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
                if (!BeginStringFormulaCell.TryCopyTo(bytes, ref written)) return false;
                buffer.Advance(written);
                return true;
            }

            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"str\"><f>"u8.TryCopyTo(bytes, ref written)) return false;

            buffer.Advance(written);
            return true;
        }

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginStyledStringCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"inlineStr\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool WriteStartElement(CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginStringCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"inlineStr\"><is><t>"u8.TryCopyTo(bytes, ref written)) return false;
        }

        buffer.Advance(written);
        return true;
    }

    public override bool WriteStartElement(StyleId styleId, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginStyledStringCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\" t=\"inlineStr\" s=\""u8.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(styleId.Id, bytes, ref written)) return false;
        if (!EndStyleBeginInlineString.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool CanWriteValuePieceByPiece(in DataCell cell) => true;

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        return buffer.WriteLongString(cell.StringValue, ref valueIndex);
    }
}
