using SpreadCheetah.CellWriters;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Time;

internal sealed class NullDateTimeCellValueWriter : NullValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.DateTimeId;

    public override bool TryWriteCell(in DataCell cell, CellWriterState state)
    {
        var defaultStyleId = state.DefaultStyling?.DateTimeStyleId;
        return defaultStyleId is { } styleId
            ? TryWriteCell(styleId, state.Buffer)
            : TryWriteCell(state.Buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = styleId?.DateTimeId ?? state.DefaultStyling?.DateTimeStyleId;
        return TryWriteCell(formulaText, actualStyleId, state.Buffer);
    }

    public override bool TryWriteCellWithReference(in DataCell cell, CellWriterState state)
    {
        var defaultStyleId = state.DefaultStyling?.DateTimeStyleId;
        return defaultStyleId is { } styleId
            ? TryWriteCellWithReference(styleId, state)
            : TryWriteCellWithReference(state);
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = styleId?.DateTimeId ?? state.DefaultStyling?.DateTimeStyleId;
        return TryWriteCellWithReference(formulaText, actualStyleId, state);
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
}
