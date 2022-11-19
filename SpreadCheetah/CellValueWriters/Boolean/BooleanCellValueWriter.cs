using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Boolean;

internal abstract class BooleanCellValueWriter : CellValueWriter
{
    private static ReadOnlySpan<byte> BeginBooleanFormulaCell => "<c t=\"b\"><f>"u8;

    protected abstract bool TryWriteCell(SpreadsheetBuffer buffer);
    protected abstract bool TryWriteEndStyleValue(Span<byte> bytes, out int bytesWritten);
    protected abstract bool TryWriteEndFormulaValue(Span<byte> bytes, out int bytesWritten);

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(buffer);
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(styleId, buffer);
    }

    private bool TryWriteCell(StyleId styleId, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();
        var part1 = StyledCellHelper.BeginStyledBooleanCell.Length;

        if (StyledCellHelper.BeginStyledBooleanCell.TryCopyTo(bytes)
            && Utf8Formatter.TryFormat(styleId.Id, bytes.Slice(part1), out var part2)
            && TryWriteEndStyleValue(bytes.Slice(part1 + part2), out var part3))
        {
            buffer.Advance(part1 + part2 + part3);
            return true;
        }

        return false;
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();

        if (TryWriteFormulaCellStart(styleId, bytes, out var part1)
            && Utf8Helper.TryGetBytes(formulaText.AsSpan(), bytes.Slice(part1), out var part2)
            && TryWriteEndFormulaValue(bytes.Slice(part1 + part2), out var part3))
        {
            buffer.Advance(part1 + part2 + part3);
            return true;
        }

        return false;
    }

    private static bool TryWriteFormulaCellStart(StyleId? styleId, Span<byte> bytes, out int bytesWritten)
    {
        if (styleId is null)
        {
            if (BeginBooleanFormulaCell.TryCopyTo(bytes))
            {
                bytesWritten = BeginBooleanFormulaCell.Length;
                return true;
            }

            bytesWritten = 0;
            return false;
        }

        var part1 = StyledCellHelper.BeginStyledBooleanCell.Length;
        var part3 = FormulaCellHelper.EndStyleBeginFormula.Length;
        if (StyledCellHelper.BeginStyledBooleanCell.TryCopyTo(bytes)
            && Utf8Formatter.TryFormat(styleId.Id, bytes.Slice(part1), out var part2)
            && FormulaCellHelper.EndStyleBeginFormula.TryCopyTo(bytes.Slice(part1 + part2)))
        {
            bytesWritten = part1 + part2 + part3;
            return true;
        }

        bytesWritten = 0;
        return false;
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer) => true;

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        if (cell.Formula is null)
            return true;

        var bytes = buffer.GetSpan();
        if (TryWriteEndFormulaValue(bytes, out var bytesWritten))
        {
            buffer.Advance(bytesWritten);
            return true;
        }

        return false;
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        if (styleId is null)
        {
            buffer.Advance(SpanHelper.GetBytes(BeginBooleanFormulaCell, buffer.GetSpan()));
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

    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer) => TryWriteCell(styleId, buffer);

    /// <summary>
    /// Returns false because the value is written together with the end element in <see cref="TryWriteEndElement(in Cell, SpreadsheetBuffer)"/>.
    /// </summary>
    public override bool CanWriteValuePieceByPiece(in DataCell cell) => false;
    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex) => true;

    public override bool Equals(in CellValue value, in CellValue other) => true;
    public override int GetHashCodeFor(in CellValue value) => 0;
}
