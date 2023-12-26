using SpreadCheetah.CellValueWriters.Number;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Buffers.Text;

namespace SpreadCheetah.CellValueWriters.Time;

internal sealed class DateTimeCellValueWriter : NumberCellValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.DateTimeId;

    protected override bool TryWriteValue(in DataCell cell, Span<byte> destination, out int bytesWritten)
    {
        return Utf8Formatter.TryFormat(cell.NumberValue.DoubleValue, destination, out bytesWritten);
    }

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return defaultStyling?.DateTimeStyleId is { } styleId
            ? TryWriteDateTimeCell(cell, styleId, buffer)
            : buffer.TryWrite($"{BeginDataCell}{cell.NumberValue.DoubleValue}{EndDefaultCell}");
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
                $"{StyledCellHelper.BeginStyledNumberCell}{style}{FormulaCellHelper.EndStyleBeginFormula}" +
                $"{formulaText}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{cachedValue.NumberValue.DoubleValue}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return buffer.TryWrite(
            $"{FormulaCellHelper.BeginNumberFormulaCell}" +
            $"{formulaText}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{cachedValue.NumberValue.DoubleValue}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    private static bool TryWriteDateTimeCell(in DataCell cell, int styleId, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite(
            $"{StyledCellHelper.BeginStyledNumberCell}{styleId}{EndStyleBeginValue}" +
            $"{cell.NumberValue.DoubleValue}" +
            $"{EndDefaultCell}");
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        if (defaultStyling?.DateTimeStyleId is { } styleId)
            return TryWriteDateTimeCellWithReference(cell, styleId, state);

        return state.Buffer.TryWrite(
            $"{state}{EndStyleBeginValue}" +
            $"{cell.NumberValue.DoubleValue}" +
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
                $"{state}{StyledCellHelper.EndReferenceBeginStyleId}{style}{FormulaCellHelper.EndStyleBeginFormula}" +
                $"{formulaText}" +
                $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
                $"{cachedValue.NumberValue.DoubleValue}" +
                $"{FormulaCellHelper.EndCachedValueEndCell}");
        }

        return state.Buffer.TryWrite(
            $"{state}{FormulaCellHelper.EndStyleBeginFormula}" +
            $"{formulaText}" +
            $"{FormulaCellHelper.EndFormulaBeginCachedValue}" +
            $"{cachedValue.NumberValue.DoubleValue}" +
            $"{FormulaCellHelper.EndCachedValueEndCell}");
    }

    private static bool TryWriteDateTimeCellWithReference(in DataCell cell, int styleId, CellWriterState state)
    {
        return state.Buffer.TryWrite(
            $"{state}{StyledCellHelper.EndReferenceBeginStyleId}{styleId}{EndStyleBeginValue}" +
            $"{cell.NumberValue.DoubleValue}" +
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
}
