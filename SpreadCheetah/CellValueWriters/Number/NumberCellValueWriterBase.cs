using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Number;

internal abstract class NumberCellValueWriterBase : CellValueWriter
{
    protected abstract int GetStyleId(StyleId styleId);
    protected abstract bool TryWriteValue(in DataCell cell, Span<byte> destination, out int bytesWritten);

    private static ReadOnlySpan<byte> BeginDataCell => "<c><v>"u8;
    private static ReadOnlySpan<byte> EndStyleBeginValue => "\"><v>"u8;
    private static ReadOnlySpan<byte> EndDefaultCell => "</v></c>"u8;

    protected bool TryWriteCell(in DataCell cell, CellWriterState state)
    {
        var buffer = state.Buffer;
        var bytes = buffer.GetSpan();
        var written = 0;

        if (!state.WriteCellReferenceAttributes)
        {
            if (!BeginDataCell.TryCopyTo(bytes, ref written)) return false;
        }
        else
        {
            if (!TryWriteCellStartWithReference(state, bytes, ref written)) return false;
            if (!"\"><v>"u8.TryCopyTo(bytes, ref written)) return false;
        }

        if (!TryWriteValue(cell, bytes.Slice(written), out var valueLength)) return false;
        written += valueLength;

        if (!EndDefaultCell.TryCopyTo(bytes, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    protected bool TryWriteCell(in DataCell cell, int styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var part1 = StyledCellHelper.BeginStyledNumberCell.Length;
        var part3 = EndStyleBeginValue.Length;
        var part5 = EndDefaultCell.Length;

        if (StyledCellHelper.BeginStyledNumberCell.TryCopyTo(bytes)
            && Utf8Formatter.TryFormat(styleId, bytes.Slice(part1), out var part2)
            && EndStyleBeginValue.TryCopyTo(bytes.Slice(part1 + part2))
            && TryWriteValue(cell, bytes.Slice(part1 + part2 + part3), out var part4)
            && EndDefaultCell.TryCopyTo(bytes.Slice(part1 + part2 + part3 + part4)))
        {
            buffer.Advance(part1 + part2 + part3 + part4 + part5);
            return true;
        }

        return false;
    }

    protected bool TryWriteCell(string formulaText, in DataCell cachedValue, int? styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var part3 = FormulaCellHelper.EndFormulaBeginCachedValue.Length;
        var part5 = FormulaCellHelper.EndCachedValueEndCell.Length;

        if (TryWriteFormulaCellStart(styleId, bytes, out var part1)
            && Utf8Helper.TryGetBytes(formulaText, bytes.Slice(part1), out var part2)
            && FormulaCellHelper.EndFormulaBeginCachedValue.TryCopyTo(bytes.Slice(part1 + part2))
            && TryWriteValue(cachedValue, bytes.Slice(part1 + part2 + part3), out var part4)
            && FormulaCellHelper.EndCachedValueEndCell.TryCopyTo(bytes.Slice(part1 + part2 + part3 + part4)))
        {
            buffer.Advance(part1 + part2 + part3 + part4 + part5);
            return true;
        }

        return false;
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(cell, GetStyleId(styleId), buffer);
    }

    public static bool TryWriteFormulaCellStart(int? styleId, Span<byte> bytes, out int bytesWritten)
    {
        if (styleId is null)
        {
            if (FormulaCellHelper.BeginNumberFormulaCell.TryCopyTo(bytes))
            {
                bytesWritten = FormulaCellHelper.BeginNumberFormulaCell.Length;
                return true;
            }

            bytesWritten = 0;
            return false;
        }

        var part1 = StyledCellHelper.BeginStyledNumberCell.Length;
        var part3 = FormulaCellHelper.EndStyleBeginFormula.Length;
        if (StyledCellHelper.BeginStyledNumberCell.TryCopyTo(bytes)
            && Utf8Formatter.TryFormat(styleId.Value, bytes.Slice(part1), out var part2)
            && FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes.Slice(part1 + part2)))
        {
            bytesWritten = part1 + part2 + part3;
            return true;
        }

        bytesWritten = 0;
        return false;
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer)
    {
        buffer.Advance(SpanHelper.GetBytes(BeginDataCell, buffer.GetSpan()));
        return true;
    }

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(GetStyleId(styleId), bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(EndStyleBeginValue, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
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

    public override bool CanWriteValuePieceByPiece(in DataCell cell) => true;

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        var bytes = buffer.GetSpan();
        if (!TryWriteValue(cell, bytes, out var bytesWritten))
            return false;

        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        if (!EndDefaultCell.TryCopyTo(bytes))
            return false;

        buffer.Advance(EndDefaultCell.Length);
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
}
