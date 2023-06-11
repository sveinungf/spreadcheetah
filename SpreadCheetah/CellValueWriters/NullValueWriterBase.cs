using SpreadCheetah.CellValueWriters.Number;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using System.Buffers.Text;

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

    protected static bool TryWriteCell(int styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();

        if (StyledCellHelper.BeginStyledNumberCell.TryCopyTo(bytes)
            && Utf8Formatter.TryFormat(styleId, bytes.Slice(StyledCellHelper.BeginStyledNumberCell.Length), out var valueLength)
            && EndStyleNullValue.TryCopyTo(bytes.Slice(StyledCellHelper.BeginStyledNumberCell.Length + valueLength)))
        {
            buffer.Advance(StyledCellHelper.BeginStyledNumberCell.Length + EndStyleNullValue.Length + valueLength);
            return true;
        }

        return false;
    }

    protected static bool TryWriteCell(string formulaText, int? styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var part3 = EndFormulaEndCell.Length;

        if (NumberCellValueWriterBase.TryWriteFormulaCellStart(styleId, bytes, out var part1)
            && Utf8Helper.TryGetBytes(formulaText, bytes.Slice(part1), out var part2)
            && EndFormulaEndCell.TryCopyTo(bytes.Slice(part1 + part2)))
        {
            buffer.Advance(part1 + part2 + part3);
            return true;
        }

        return false;
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(GetStyleId(styleId), buffer);
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

    protected static bool WriteFormulaStartElement(int? styleId, SpreadsheetBuffer buffer)
    {
        if (styleId is null)
        {
            buffer.Advance(SpanHelper.GetBytes(FormulaCellHelper.BeginNumberFormulaCell, buffer.GetSpan()));
            return true;
        }

        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId.Value, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer)
        => TryWriteCell(new CellWriterState(buffer, false)); // TODO: Temporary workaround

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer) => TryWriteCell(GetStyleId(styleId), buffer);

    /// <summary>
    /// Returns false because there is no value to write for 'null'.
    /// </summary>
    public override bool CanWriteValuePieceByPiece(in DataCell cell) => false;
    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex) => true;

    public override bool Equals(in CellValue value, in CellValue other) => true;
    public override int GetHashCodeFor(in CellValue value) => 0;
}
