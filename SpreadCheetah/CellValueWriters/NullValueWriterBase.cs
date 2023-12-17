using SpreadCheetah.CellValueWriters.Number;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters;

internal abstract class NullValueWriterBase : CellValueWriter
{
    protected abstract int GetStyleId(StyleId styleId);

    private static ReadOnlySpan<byte> NullDataCell => "<c/>"u8;
    private static ReadOnlySpan<byte> EndStyleNullValue => "\"/>"u8;
    private static ReadOnlySpan<byte> EndFormulaEndCell => "</f></c>"u8;

    protected static bool TryWriteCell(CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!state.WriteCellReferenceAttributes)
        {
            if (!NullDataCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\"/>"u8.TryCopyTo(bytes, ref written)) return false;
        }

        buffer.Advance(written);
        return true;
    }

    protected static bool TryWriteCell(int styleId, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!state.WriteCellReferenceAttributes)
        {
            if (!StyledCellHelper.BeginStyledNumberCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!StyledCellHelper.EndReferenceBeginStyleId.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(styleId, bytes, ref written)) return false;
        if (!EndStyleNullValue.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    protected static bool TryWriteCell(int styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!StyledCellHelper.BeginStyledNumberCell.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(styleId, bytes, ref written)) return false;
        if (!EndStyleNullValue.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    protected static bool TryWriteCell(string formulaText, int? styleId, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();

        if (!NumberCellValueWriterBase.TryWriteFormulaCellStart(styleId, state, bytes, out var written)) return false;
        if (!SpanHelper.TryWrite(formulaText, bytes, ref written)) return false;
        if (!EndFormulaEndCell.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return TryWriteCell(GetStyleId(styleId), state);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(GetStyleId(styleId), buffer);
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return TryWriteCellWithReference(GetStyleId(styleId), state);
    }

    protected static bool TryWriteCellWithReference(int styleId, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
        if (!StyledCellHelper.EndReferenceBeginStyleId.TryCopyTo(bytes, ref written)) return false;
        if (!SpanHelper.TryWrite(styleId, bytes, ref written)) return false;
        if (!EndStyleNullValue.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer) => true;

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        if (cell.Formula is null)
            return true;

        var bytes = buffer.GetSpan();
        if (EndFormulaEndCell.TryCopyTo(bytes))
        {
            buffer.Advance(EndFormulaEndCell.Length);
            return true;
        }

        return false;
    }

    protected static bool WriteFormulaStartElement(int? styleId, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (styleId is null)
        {
            if (!state.WriteCellReferenceAttributes)
            {
                if (!FormulaCellHelper.BeginNumberFormulaCell.TryCopyTo(bytes, ref written)) return false;
                buffer.Advance(written);
                return true;
            }

            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\"><f>"u8.TryCopyTo(bytes, ref written)) return false;

            buffer.Advance(written);
            return true;
        }

        if (!state.WriteCellReferenceAttributes)
        {
            if (!StyledCellHelper.BeginStyledNumberCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!StyledCellHelper.EndReferenceBeginStyleId.TryCopyTo(bytes, ref written)) return false;
        }

        if (!SpanHelper.TryWrite(styleId.Value, bytes, ref written)) return false;
        if (!FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    public override bool WriteStartElement(CellWriterState state) => TryWriteCell(state);

    public override bool WriteStartElement(StyleId styleId, CellWriterState state) => TryWriteCell(GetStyleId(styleId), state);

    /// <summary>
    /// Returns false because there is no value to write for 'null'.
    /// </summary>
    public override bool CanWriteValuePieceByPiece(in DataCell cell) => false;
    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex) => true;
}
