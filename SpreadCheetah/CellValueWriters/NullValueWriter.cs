using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters;

internal sealed class NullValueWriter : CellValueWriter
{
    private static readonly int DataCellElementLength =
        DataCellHelper.NullCell.Length;

    private static readonly int StyledCellElementLength =
        StyledCellHelper.BeginStyledNumberCell.Length +
        SpreadsheetConstants.StyleIdMaxDigits +
        StyledCellHelper.EndStyleNullValue.Length;

    private static readonly int FormulaCellElementLength =
        StyledCellHelper.BeginStyledNumberCell.Length +
        SpreadsheetConstants.StyleIdMaxDigits +
        FormulaCellHelper.EndStyleBeginFormula.Length +
        FormulaCellHelper.EndFormulaEndCell.Length;

    public override bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer) => GetBytes(buffer);
    public override bool GetBytes(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer) => GetBytes(styleId, buffer);

    private static bool GetBytes(SpreadsheetBuffer buffer)
    {
        buffer.Advance(SpanHelper.GetBytes(DataCellHelper.NullCell, buffer.GetNextSpan()));
        return true;
    }

    private static bool GetBytes(StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetNextSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleNullValue, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool GetBytes(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetNextSpan();
        int bytesWritten;

        if (styleId is null)
        {
            bytesWritten = SpanHelper.GetBytes(FormulaCellHelper.BeginNumberFormulaCell, bytes);
        }
        else
        {
            bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        }

        bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), false);
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndFormulaEndCell, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool TryWriteCell(in DataCell cell, SpreadsheetBuffer buffer, out int bytesNeeded)
    {
        bytesNeeded = DataCellElementLength;
        var remaining = buffer.GetRemainingBuffer();
        return bytesNeeded <= remaining && GetBytes(cell, buffer);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer, out int bytesNeeded)
    {
        bytesNeeded = StyledCellElementLength;
        var remaining = buffer.GetRemainingBuffer();
        return bytesNeeded <= remaining && GetBytes(cell, styleId, buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer, out int bytesNeeded)
    {
        // Try with approximate formula text length
        bytesNeeded = FormulaCellElementLength + formulaText.Length * Utf8Helper.MaxBytePerChar;
        var remaining = buffer.GetRemainingBuffer();
        if (bytesNeeded <= remaining)
            return GetBytes(formulaText, cachedValue, styleId, buffer);

        // Try with more accurate length
        bytesNeeded = FormulaCellElementLength + Utf8Helper.GetByteCount(formulaText);
        return bytesNeeded <= remaining && GetBytes(formulaText, cachedValue, styleId, buffer);
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer) => true;

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        if (cell.Formula is null)
            return true;

        var cellEnd = FormulaCellHelper.EndFormulaEndCell;
        if (cellEnd.Length > buffer.GetRemainingBuffer())
            return false;

        buffer.Advance(SpanHelper.GetBytes(cellEnd, buffer.GetNextSpan()));
        return true;
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, SpreadsheetBuffer buffer)
    {
        if (styleId is null)
        {
            buffer.Advance(SpanHelper.GetBytes(FormulaCellHelper.BeginNumberFormulaCell, buffer.GetNextSpan()));
            return true;
        }

        var bytes = buffer.GetNextSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer) => GetBytes(buffer);

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer) => GetBytes(styleId, buffer);

    public override bool Equals(in CellValue value, in CellValue other) => true;
    public override int GetHashCodeFor(in CellValue value) => 0;
}
