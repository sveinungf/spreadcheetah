using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters;

internal abstract class NullValueWriterBase : CellValueWriter
{
    protected abstract int GetStyleId(StyleId styleId);

    private static ReadOnlySpan<byte> NullDataCell => "<c/>"u8;
    private static ReadOnlySpan<byte> EndQuoteNullValue => "\"/>"u8;
    private static ReadOnlySpan<byte> EndFormulaEndCell => "</f></c>"u8;

    protected static bool TryWriteCell(SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{NullDataCell}");
    }

    protected static bool TryWriteCell(int styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{StyledCellHelper.BeginStyledNumberCell}{styleId}{EndQuoteNullValue}");
    }

    protected static bool TryWriteCell(string formulaText, int? styleId, SpreadsheetBuffer buffer)
    {
        if (styleId is { } style)
        {
            return buffer.TryWrite(
                $"{StyledCellHelper.BeginStyledNumberCell}{style}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{new RawString(formulaText)}" +
                $"{EndFormulaEndCell}");
        }

        return buffer.TryWrite(
            $"{FormulaCellHelper.BeginNumberFormulaCell}" +
            $"{new RawString(formulaText)}" +
            $"{EndFormulaEndCell}");
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(GetStyleId(styleId), buffer);
    }

    protected static bool TryWriteCellWithReference(CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{EndQuoteNullValue}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return TryWriteCellWithReference(GetStyleId(styleId), state);
    }

    protected static bool TryWriteCellWithReference(int styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite($"{state}{StyledCellHelper.EndReferenceBeginStyleId}{styleId}{EndQuoteNullValue}");
    }

    protected static bool TryWriteCellWithReference(string formulaText, int? styleId, CellWriterState state)
    {
        if (styleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{state}{StyledCellHelper.EndReferenceBeginStyleId}{style}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{EndFormulaEndCell}");
        }

        return state.Buffer.TryWrite(
            $"{state}{FormulaCellHelper.EndQuoteBeginFormula}" +
            $"{formulaText}" +
            $"{EndFormulaEndCell}");
    }

    public override bool TryWriteEndElement(SpreadsheetBuffer buffer) => true;

    public override bool TryWriteEndElement(in Cell cell, SpreadsheetBuffer buffer)
    {
        return cell.Formula is null || buffer.TryWrite($"{EndFormulaEndCell}");
    }

    protected static bool WriteFormulaStartElement(int? styleId, SpreadsheetBuffer buffer)
    {
        return styleId is { } style
            ? buffer.TryWrite($"{StyledCellHelper.BeginStyledNumberCell}{style}{FormulaCellHelper.EndQuoteBeginFormula}")
            : buffer.TryWrite($"{FormulaCellHelper.BeginNumberFormulaCell}");
    }

    protected static bool WriteFormulaStartElementWithReference(int? styleId, CellWriterState state)
    {
        return styleId is { } style
            ? state.Buffer.TryWrite($"{state}{StyledCellHelper.EndReferenceBeginStyleId}{style}{FormulaCellHelper.EndQuoteBeginFormula}")
            : state.Buffer.TryWrite($"{state}{FormulaCellHelper.EndQuoteBeginFormula}");
    }

    public override bool WriteStartElement(SpreadsheetBuffer buffer) => TryWriteCell(buffer);
    public override bool WriteStartElement(StyleId styleId, SpreadsheetBuffer buffer) => TryWriteCell(GetStyleId(styleId), buffer);
    public override bool WriteStartElementWithReference(CellWriterState state) => TryWriteCellWithReference(state);
    public override bool WriteStartElementWithReference(StyleId styleId, CellWriterState state) => TryWriteCellWithReference(GetStyleId(styleId), state);

    /// <summary>
    /// Returns false because there is no value to write for 'null'.
    /// </summary>
    public override bool CanWriteValuePieceByPiece(in DataCell cell) => false;
    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex) => true;
}
