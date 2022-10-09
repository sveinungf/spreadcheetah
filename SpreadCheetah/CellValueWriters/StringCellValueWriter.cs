using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters;

internal sealed class StringCellValueWriter : CellValueWriter
{
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
        var bytes = buffer.GetSpan();
        var part3 = FormulaCellHelper.EndFormulaBeginCachedValue.Length;
        var part5 = FormulaCellHelper.EndCachedValueEndCell.Length;

        if (TryWriteFormulaCellStart(styleId, bytes, out var part1)
            && Utf8Helper.TryGetBytes(formulaText.AsSpan(), bytes.Slice(part1), out var part2)
            && FormulaCellHelper.EndFormulaBeginCachedValue.TryCopyTo(bytes.Slice(part1 + part2))
            && Utf8Helper.TryGetBytes(cachedValue.StringValue!.AsSpan(), bytes.Slice(part1 + part2 + part3), out var part4)
            && FormulaCellHelper.EndCachedValueEndCell.TryCopyTo(bytes.Slice(part1 + part2 + part3 + part4)))
        {
            buffer.Advance(part1 + part2 + part3 + part4 + part5);
            return true;
        }

        return false;
    }

    private static bool TryWriteFormulaCellStart(StyleId? styleId, Span<byte> bytes, out int bytesWritten)
    {
        if (styleId is null)
        {
            if (FormulaCellHelper.BeginStringFormulaCell.TryCopyTo(bytes))
            {
                bytesWritten = FormulaCellHelper.BeginStringFormulaCell.Length;
                return true;
            }

            bytesWritten = 0;
            return false;
        }

        var part1 = FormulaCellHelper.BeginStyledStringFormulaCell.Length;
        var part3 = FormulaCellHelper.EndStyleBeginFormula.Length;
        if (FormulaCellHelper.BeginStyledStringFormulaCell.TryCopyTo(bytes)
            && Utf8Formatter.TryFormat(styleId.Id, bytes.Slice(part1), out var part2)
            && FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes.Slice(part1 + part2)))
        {
            bytesWritten = part1 + part2 + part3;
            return true;
        }

        bytesWritten = 0;
        return false;
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        if (!DataCellHelper.EndStringCell.TryCopyTo(bytes))
            return false;

        buffer.Advance(DataCellHelper.EndStringCell.Length);
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
