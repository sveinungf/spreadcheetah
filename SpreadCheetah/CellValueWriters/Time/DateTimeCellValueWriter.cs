using SpreadCheetah.CellValueWriters.Number;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Time;

internal sealed class DateTimeCellValueWriter : NumberCellValueWriterBase
{
    public override bool TryWriteCell(in DataCell cell, CellWriterState state)
    {
        return state.DefaultStyling?.DateTimeStyleId is { } styleId
            ? TryWriteDateTimeCell(cell, styleId, state.Buffer)
            : state.Buffer.TryWrite($"{BeginDataCell}{new OADate(cell.Value.StringOrPrimitive.PrimitiveValue.LongValue)}{EndDefaultCell}");
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteDateTimeCell(cell, styleId.DateTimeId, buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = styleId?.DateTimeId ?? state.DefaultStyling?.DateTimeStyleId;
        if (actualStyleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{StyledCellHelper.BeginStyledNumberCell}{style}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{new OADate(cachedValue.Value.StringOrPrimitive.PrimitiveValue.LongValue)}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return state.Buffer.TryWrite(
            $"{FormulaCellHelper.BeginNumberFormulaCell}" +
            $"{formulaText}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{new OADate(cachedValue.Value.StringOrPrimitive.PrimitiveValue.LongValue)}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    private static bool TryWriteDateTimeCell(in DataCell cell, int styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite(
            $"{StyledCellHelper.BeginStyledNumberCell}{styleId}{EndQuoteBeginValue}" +
            $"{new OADate(cell.Value.StringOrPrimitive.PrimitiveValue.LongValue)}" +
            $"{EndDefaultCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, CellWriterState state)
    {
        if (state.DefaultStyling?.DateTimeStyleId is { } styleId)
            return TryWriteDateTimeCellWithReference(cell, styleId, state);

        return state.Buffer.TryWrite(
            $"{state}{EndQuoteBeginValue}" +
            $"{new OADate(cell.Value.StringOrPrimitive.PrimitiveValue.LongValue)}" +
            $"{EndDefaultCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return TryWriteDateTimeCellWithReference(cell, styleId.DateTimeId, state);
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = styleId?.DateTimeId ?? state.DefaultStyling?.DateTimeStyleId;
        if (actualStyleId is { } style)
        {
            return state.Buffer.TryWrite(
                $"{state}{StyledCellHelper.EndReferenceBeginStyleId}{style}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{new OADate(cachedValue.Value.StringOrPrimitive.PrimitiveValue.LongValue)}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return state.Buffer.TryWrite(
            $"{state}{FormulaCellHelper.EndQuoteBeginFormula}" +
            $"{formulaText}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{new OADate(cachedValue.Value.StringOrPrimitive.PrimitiveValue.LongValue)}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    private static bool TryWriteDateTimeCellWithReference(in DataCell cell, int styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite(
            $"{state}{StyledCellHelper.EndReferenceBeginStyleId}{styleId}{EndQuoteBeginValue}" +
            $"{new OADate(cell.Value.StringOrPrimitive.PrimitiveValue.LongValue)}" +
            $"{EndDefaultCell}");
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = styleId?.DateTimeId ?? state.DefaultStyling?.DateTimeStyleId;
        return WriteFormulaStartElement(actualStyleId, state.Buffer);
    }

    public override bool WriteFormulaStartElementWithReference(StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = styleId?.DateTimeId ?? state.DefaultStyling?.DateTimeStyleId;
        return WriteFormulaStartElementWithReference(actualStyleId, state);
    }

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        return buffer.TryWrite($"{new OADate(cell.Value.StringOrPrimitive.PrimitiveValue.LongValue)}");
    }
}
