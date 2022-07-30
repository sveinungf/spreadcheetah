using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Number;

internal abstract class NumberCellValueWriter : CellValueWriter
{
    private static readonly int DataCellElementLength =
        DataCellHelper.BeginNumberCell.Length +
        DataCellHelper.EndDefaultCell.Length;

    private static readonly int StyledCellElementLength =
        StyledCellHelper.BeginStyledNumberCell.Length +
        SpreadsheetConstants.StyleIdMaxDigits +
        StyledCellHelper.EndStyleBeginValue.Length +
        DataCellHelper.EndDefaultCell.Length;

    private static readonly int FormulaCellElementLength =
        StyledCellHelper.BeginStyledNumberCell.Length +
        SpreadsheetConstants.StyleIdMaxDigits +
        FormulaCellHelper.EndStyleBeginFormula.Length +
        FormulaCellHelper.EndFormulaBeginCachedValue.Length +
        FormulaCellHelper.EndCachedValueEndCell.Length;

    protected abstract int MaxNumberLength { get; }
    protected abstract int GetValueBytes(in DataCell cell, Span<byte> destination);

    private bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(DataCellHelper.BeginNumberCell, bytes);
        bytesWritten += GetValueBytes(cell, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(DataCellHelper.EndDefaultCell, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    private bool GetBytes(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleBeginValue, bytes.Slice(bytesWritten));
        bytesWritten += GetValueBytes(cell, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(DataCellHelper.EndDefaultCell, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    private bool GetBytes(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer)
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
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        }

        bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), false);
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndFormulaBeginCachedValue, bytes.Slice(bytesWritten));
        bytesWritten += GetValueBytes(cachedValue, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndCachedValueEndCell, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool TryWriteCell(in DataCell cell, SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;
        return DataCellElementLength + MaxNumberLength <= remaining && GetBytes(cell, buffer);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;
        return StyledCellElementLength + MaxNumberLength <= remaining && GetBytes(cell, styleId, buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;

        // Try with approximate formula text length
        var bytesNeeded = FormulaCellElementLength + MaxNumberLength + formulaText.Length * Utf8Helper.MaxBytePerChar;
        if (bytesNeeded <= remaining)
            return GetBytes(formulaText, cachedValue, styleId, buffer);

        // Try with more accurate length
        bytesNeeded = FormulaCellElementLength + MaxNumberLength + Utf8Helper.GetByteCount(formulaText);
        return bytesNeeded <= remaining && GetBytes(formulaText, cachedValue, styleId, buffer);
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer)
    {
        buffer.Advance(SpanHelper.GetBytes(DataCellHelper.BeginNumberCell, buffer.GetSpan()));
        return true;
    }

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleBeginValue, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, SpreadsheetBuffer buffer)
    {
        return DoWriteFormulaStartElement(styleId, buffer);
    }

    public static bool DoWriteFormulaStartElement(StyleId? styleId, SpreadsheetBuffer buffer)
    {
        if (styleId is null)
        {
            buffer.Advance(SpanHelper.GetBytes(FormulaCellHelper.BeginNumberFormulaCell, buffer.GetSpan()));
            return true;
        }

        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer)
    {
        var cellEnd = DataCellHelper.EndDefaultCell;
        var bytes = buffer.GetSpan();
        if (cellEnd.Length >= bytes.Length)
            return false;

        buffer.Advance(SpanHelper.GetBytes(cellEnd, bytes));
        return true;
    }

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        if (cell.Formula is null)
            return TryWriteEndElement(buffer);

        var cellEnd = FormulaCellHelper.EndCachedValueEndCell;
        if (cellEnd.Length > buffer.FreeCapacity)
            return false;

        buffer.Advance(SpanHelper.GetBytes(cellEnd, buffer.GetSpan()));
        return true;
    }
}
