using SpreadCheetah.CellValueWriters.Number;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Time;

internal sealed class DateTimeCellValueWriter : NumberCellValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.DateTimeId;

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return defaultStyling?.DateTimeStyleId is { } styleId
            ? TryWriteDateTimeCell(cell, styleId, buffer)
            : buffer.TryWrite($"{BeginDataCell}{new OADate(cell.Value.StringOrPrimitive.PrimitiveValue.LongValue)}{EndDefaultCell}");
    }

    public override bool TryWriteCell(in DataCell cell, StyleId styleId, SpreadsheetBuffer buffer)
    {
        return TryWriteDateTimeCell(cell, styleId.DateTimeId, buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var actualStyleId = styleId?.DateTimeId ?? defaultStyling?.DateTimeStyleId;
        if (actualStyleId is { } style)
        {
            return buffer.TryWrite(
                $"{StyledCellHelper.BeginStyledNumberCell}{style}{FormulaCellHelper.EndQuoteBeginFormula}" +
                $"{formulaText}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{new OADate(cachedValue.Value.StringOrPrimitive.PrimitiveValue.LongValue)}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return buffer.TryWrite(
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

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        if (defaultStyling?.DateTimeStyleId is { } styleId)
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

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var actualStyleId = styleId?.DateTimeId ?? defaultStyling?.DateTimeStyleId;
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

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var actualStyleId = styleId?.DateTimeId ?? defaultStyling?.DateTimeStyleId;
        return WriteFormulaStartElement(actualStyleId, buffer);
    }

    public override bool WriteFormulaStartElementWithReference(StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var actualStyleId = styleId?.DateTimeId ?? defaultStyling?.DateTimeStyleId;
        return WriteFormulaStartElementWithReference(actualStyleId, state);
    }

    public override bool WriteValuePieceByPiece(in DataCell cell, SpreadsheetBuffer buffer, ref int valueIndex)
    {
        return buffer.TryWrite($"{new OADate(cell.Value.StringOrPrimitive.PrimitiveValue.LongValue)}");
    }
}
