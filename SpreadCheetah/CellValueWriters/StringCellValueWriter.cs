using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters;

internal sealed class StringCellValueWriter : CellValueWriter
{
    private static readonly int FormulaCellElementLength =
        FormulaCellHelper.BeginStyledStringFormulaCell.Length +
        SpreadsheetConstants.StyleIdMaxDigits +
        FormulaCellHelper.EndStyleBeginFormula.Length +
        FormulaCellHelper.EndFormulaBeginCachedValue.Length +
        FormulaCellHelper.EndCachedValueEndCell.Length;

    private static bool GetBytes(string formulaText, in DataCell cachedValue, StyleId? styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        int bytesWritten;

        if (styleId is null)
        {
            bytesWritten = SpanHelper.GetBytes(FormulaCellHelper.BeginStringFormulaCell, bytes);
        }
        else
        {
            bytesWritten = SpanHelper.GetBytes(FormulaCellHelper.BeginStyledStringFormulaCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        }

        bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), false);
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndFormulaBeginCachedValue, bytes.Slice(bytesWritten));
        bytesWritten += Utf8Helper.GetBytes(cachedValue.StringValue!, bytes.Slice(bytesWritten), false);
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndCachedValueEndCell, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();

        if (DataCellHelper.BeginStringCell.TryCopyTo(bytes)
            && Utf8Helper.TryGetBytes(cell.StringValue!.AsSpan(), bytes.Slice(DataCellHelper.BeginStringCell.Length), out var valueLength)
            && DataCellHelper.EndStringCell.TryCopyTo(bytes.Slice(DataCellHelper.BeginStringCell.Length + valueLength)))
        {
            buffer.Advance(DataCellHelper.BeginStringCell.Length + DataCellHelper.EndStringCell.Length + valueLength);
            return true;
        }

        return false;
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var part1 = StyledCellHelper.BeginStyledStringCell.Length;
        var part3 = StyledCellHelper.EndStyleBeginInlineString.Length;
        var part5 = DataCellHelper.EndStringCell.Length;

        if (StyledCellHelper.BeginStyledStringCell.TryCopyTo(bytes)
            && Utf8Formatter.TryFormat(styleId.Id, bytes.Slice(part1), out var part2)
            && StyledCellHelper.EndStyleBeginInlineString.TryCopyTo(bytes.Slice(part1 + part2))
            && Utf8Helper.TryGetBytes(cell.StringValue!.AsSpan(), bytes.Slice(part1 + part2 + part3), out var part4)
            && DataCellHelper.EndStringCell.TryCopyTo(bytes.Slice(part1 + part2 + part3 + part4)))
        {
            buffer.Advance(part1 + part2 + part3 + part4 + part5);
            return true;
        }

        return false;
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;

        // Try with an approximate cell value and formula text length
        var bytesNeeded = FormulaCellElementLength + (formulaText.Length + cachedValue.StringValue!.Length) * Utf8Helper.MaxBytePerChar;
        if (bytesNeeded <= remaining)
            return GetBytes(formulaText, cachedValue, styleId, buffer);

        // Try with more accurate length
        bytesNeeded = FormulaCellElementLength + Utf8Helper.GetByteCount(formulaText) + Utf8Helper.GetByteCount(cachedValue.StringValue);
        return bytesNeeded <= remaining && GetBytes(formulaText, cachedValue, styleId, buffer);
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer)
    {
        var cellEnd = DataCellHelper.EndStringCell;
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

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        if (styleId is null)
        {
            buffer.Advance(SpanHelper.GetBytes(FormulaCellHelper.BeginStringFormulaCell, buffer.GetSpan()));
            return true;
        }

        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledStringCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer)
    {
        buffer.Advance(SpanHelper.GetBytes(DataCellHelper.BeginStringCell, buffer.GetSpan()));
        return true;
    }

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledStringCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleBeginInlineString, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool CanWriteValuePieceByPiece(in DataCell cell) => true;

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        return buffer.WriteLongString(cell.StringValue.AsSpan(), ref valueIndex);
    }

    public override bool Equals(in CellValue value, in CellValue other) => true;
    public override int GetHashCodeFor(in CellValue value) => 0;
}
