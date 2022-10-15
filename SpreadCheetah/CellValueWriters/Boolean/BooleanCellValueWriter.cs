using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal abstract class BooleanCellValueWriter : CellValueWriter
{
    private static readonly int StyledCellElementLength =
        StyledCellHelper.BeginStyledBooleanCell.Length +
        SpreadsheetConstants.StyleIdMaxDigits +
        StyledCellHelper.EndStyleTrueBooleanValue.Length;

    private static readonly int FormulaCellElementLength =
        StyledCellHelper.BeginStyledBooleanCell.Length +
        SpreadsheetConstants.StyleIdMaxDigits +
        FormulaCellHelper.EndStyleBeginFormula.Length +
        FormulaCellHelper.EndFormulaTrueBooleanValue.Length;

    protected abstract ReadOnlySpan<byte> DataCellBytes();
    protected abstract ReadOnlySpan<byte> EndFormulaValueBytes();
    protected abstract ReadOnlySpan<byte> EndStyleValueBytes();
    protected abstract bool TryWriteCell(SpreadsheetBuffer buffer);

    private bool GetBytes(StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledBooleanCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(EndStyleValueBytes(), bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    private bool GetBytes(string formulaText, StyleId? styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        int bytesWritten;

        if (styleId is null)
        {
            bytesWritten = SpanHelper.GetBytes(FormulaCellHelper.BeginBooleanFormulaCell, bytes);
        }
        else
        {
            bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledBooleanCell, bytes);
            bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        }

        bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), false);
        bytesWritten += SpanHelper.GetBytes(EndFormulaValueBytes(), bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(buffer);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;
        return StyledCellElementLength <= remaining && GetBytes(styleId, buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        // Try with approximate formula text length
        var bytesNeeded = FormulaCellElementLength + formulaText.Length * Utf8Helper.MaxBytePerChar;
        var remaining = buffer.FreeCapacity;
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

        var cellEnd = EndFormulaValueBytes();
        if (cellEnd.Length > buffer.FreeCapacity)
            return false;

        buffer.Advance(SpanHelper.GetBytes(cellEnd, buffer.GetSpan()));
        return true;
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        if (styleId is null)
        {
            buffer.Advance(SpanHelper.GetBytes(FormulaCellHelper.BeginBooleanFormulaCell, buffer.GetSpan()));
            return true;
        }

        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledBooleanCell, buffer.GetSpan());
        bytesWritten += Utf8Helper.GetBytes(styleId.Id, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer) => TryWriteCell(buffer);

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer) => GetBytes(styleId, buffer);

    /// <summary>
    /// Returns false because the value is written together with the end element in <see cref="TryWriteEndElement(in Cell, SpreadsheetBuffer)"/>.
    /// </summary>
    public override bool CanWriteValuePieceByPiece(in DataCell cell) => false;
    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex) => true;

    public override bool Equals(in CellValue value, in CellValue other) => true;
    public override int GetHashCodeFor(in CellValue value) => 0;
}
