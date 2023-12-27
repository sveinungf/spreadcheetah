using SpreadCheetah.CellWriters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Time;

internal sealed class NullDateTimeCellValueWriter : NullValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.DateTimeId;

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var defaultStyleId = defaultStyling?.DateTimeStyleId;
        return defaultStyleId is { } styleId
            ? TryWriteCell(styleId, buffer)
            : TryWriteCell(buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var actualStyleId = styleId?.DateTimeId ?? defaultStyling?.DateTimeStyleId;
        return TryWriteCell(formulaText, actualStyleId, buffer);
    }

    public override bool TryWriteCellWithReference(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var defaultStyleId = defaultStyling?.DateTimeStyleId;
        return defaultStyleId is { } styleId
            ? TryWriteCellWithReference(styleId, state)
            : TryWriteCellWithReference(state);
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        var actualStyleId = styleId?.DateTimeId ?? defaultStyling?.DateTimeStyleId;
        return TryWriteCellWithReference(formulaText, actualStyleId, state);
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
