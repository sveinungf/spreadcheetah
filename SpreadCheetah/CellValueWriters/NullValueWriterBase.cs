using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters;

internal abstract class NullValueWriterBase : CellValueWriter
{
    private static readonly int StyledCellElementLength =
        StyledCellHelper.BeginStyledNumberCell.Length +
        SpreadsheetConstants.StyleIdMaxDigits +
        StyledCellHelper.EndStyleNullValue.Length;

    private static readonly int FormulaCellElementLength =
        StyledCellHelper.BeginStyledNumberCell.Length +
        SpreadsheetConstants.StyleIdMaxDigits +
        FormulaCellHelper.EndStyleBeginFormula.Length +
        FormulaCellHelper.EndFormulaEndCell.Length;

    private static int DataCellElementLength => NullDataCell().Length;
    protected abstract int GetStyleId(StyleId styleId);

    // <c/>
    private static ReadOnlySpan<byte> NullDataCell() => new[]
    {
        (byte)'<', (byte)'c', (byte)'/', (byte)'>'
    };

    private static bool GetBytes(SpreadsheetBuffer buffer)
    {
        buffer.Advance(SpanHelper.GetBytes(NullDataCell(), buffer.GetSpan()));
        return true;
    }

    private static bool GetBytes(int styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleNullValue, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    private bool GetBytes(string formulaText, StyleId? styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        int bytesWritten;

        if (styleId is null)
        {
            bytesWritten = SpanHelper.GetBytes(FormulaCellHelper.BeginNumberFormulaCell, bytes);
        }
        else
        {
            bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(GetStyleId(styleId), bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        }

        bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), false);
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndFormulaEndCell, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    protected static bool TryWriteCell(SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;
        return DataCellElementLength <= remaining && GetBytes(buffer);
    }

    protected static bool TryWriteCell(int styleId, SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;
        return StyledCellElementLength <= remaining && GetBytes(styleId, buffer);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(GetStyleId(styleId), buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;

        // Try with approximate formula text length
        var bytesNeeded = FormulaCellElementLength + formulaText.Length * Utf8Helper.MaxBytePerChar;
        if (bytesNeeded <= remaining)
            return GetBytes(formulaText, styleId, buffer);

        // Try with more accurate length
        bytesNeeded = FormulaCellElementLength + Utf8Helper.GetByteCount(formulaText);
        return bytesNeeded <= remaining && GetBytes(formulaText, styleId, buffer);
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer) => true;

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        if (cell.Formula is null)
            return true;

        var cellEnd = FormulaCellHelper.EndFormulaEndCell;
        if (cellEnd.Length > buffer.FreeCapacity)
            return false;

        buffer.Advance(SpanHelper.GetBytes(cellEnd, buffer.GetSpan()));
        return true;
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, SpreadsheetBuffer buffer)
    {
        if (styleId is null)
        {
            buffer.Advance(SpanHelper.GetBytes(FormulaCellHelper.BeginNumberFormulaCell, buffer.GetSpan()));
            return true;
        }

        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(GetStyleId(styleId), bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer) => GetBytes(buffer);

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer) => GetBytes(GetStyleId(styleId), buffer);

    /// <summary>
    /// Returns false because there is no value to write for 'null'.
    /// </summary>
    public override bool CanWriteValuePieceByPiece(in DataCell cell) => false;
    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex) => true;

    public override bool Equals(in CellValue value, in CellValue other) => true;
    public override int GetHashCodeFor(in CellValue value) => 0;
}
