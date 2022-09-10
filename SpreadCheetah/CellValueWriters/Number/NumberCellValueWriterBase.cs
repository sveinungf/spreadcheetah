using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Number;

internal abstract class NumberCellValueWriterBase : CellValueWriter
{
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

    private static int DataCellElementLength =>
        BeginDataCell().Length +
        DataCellHelper.EndDefaultCell.Length;

    protected abstract int MaxNumberLength { get; }
    protected abstract int GetStyleId(StyleId styleId);
    protected abstract int GetValueBytes(in DataCell cell, Span<byte> destination);

    // <c><v>
    private static ReadOnlySpan<byte> BeginDataCell() => new[]
    {
        (byte)'<', (byte)'c', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
    };

    private bool GetBytes(in DataCell cell, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(BeginDataCell(), bytes);
        bytesWritten += GetValueBytes(cell, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(DataCellHelper.EndDefaultCell, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    protected bool GetBytes(in DataCell cell, int styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(styleId, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleBeginValue, bytes.Slice(bytesWritten));
        bytesWritten += GetValueBytes(cell, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(DataCellHelper.EndDefaultCell, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    private bool GetBytes(string formulaText, in DataCell cachedValue, int? styleId, SpreadsheetBuffer buffer)
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
            bytesWritten += Utf8Helper.GetBytes(styleId.Value, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndStyleBeginFormula, bytes.Slice(bytesWritten));
        }

        bytesWritten += Utf8Helper.GetBytes(formulaText, bytes.Slice(bytesWritten), false);
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndFormulaBeginCachedValue, bytes.Slice(bytesWritten));
        bytesWritten += GetValueBytes(cachedValue, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(FormulaCellHelper.EndCachedValueEndCell, bytes.Slice(bytesWritten));
        buffer.Advance(bytesWritten);
        return true;
    }

    protected bool TryWriteCell(in DataCell cell, SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;
        return DataCellElementLength + MaxNumberLength <= remaining && GetBytes(cell, buffer);
    }

    protected bool TryWriteCell(in DataCell cell, int styleId, SpreadsheetBuffer buffer)
    {
        var remaining = buffer.FreeCapacity;
        return StyledCellElementLength + MaxNumberLength <= remaining && GetBytes(cell, styleId, buffer);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(cell, GetStyleId(styleId), buffer);
    }

    protected bool TryWriteCell(string formulaText, in DataCell cachedValue, int? styleId, SpreadsheetBuffer buffer)
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
        buffer.Advance(SpanHelper.GetBytes(BeginDataCell(), buffer.GetSpan()));
        return true;
    }

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var bytesWritten = SpanHelper.GetBytes(StyledCellHelper.BeginStyledNumberCell, bytes);
        bytesWritten += Utf8Helper.GetBytes(GetStyleId(styleId), bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(StyledCellHelper.EndStyleBeginValue, bytes.Slice(bytesWritten));
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
        if (MaxNumberLength > buffer.FreeCapacity) return false;
        buffer.Advance(GetValueBytes(cell, buffer.GetSpan()));
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
