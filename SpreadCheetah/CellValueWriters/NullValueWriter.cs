using SpreadCheetah.CellWriters;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters;

internal sealed class NullValueWriter : NullValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.Id;

    public override bool TryWriteCell(in DataCell cell, CellWriterState state)
    {
        return TryWriteCell(state.Buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, CellWriterState state)
    {
        return TryWriteCell(formulaText, styleId?.Id, state.Buffer);
    }

    public override bool TryWriteCellWithReference(in DataCell cell, CellWriterState state)
    {
        return TryWriteCellWithReference(state);
    }

    public override bool TryWriteCellWithReference(string formulaText, in DataCell cachedValue, StyleId? styleId, CellWriterState state)
    {
        return TryWriteCellWithReference(formulaText, styleId?.Id, state);
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, CellWriterState state)
    {
        return WriteFormulaStartElement(styleId?.Id, state.Buffer);
    }

    public override bool WriteFormulaStartElementWithReference(StyleId? styleId, CellWriterState state)
    {
        return WriteFormulaStartElementWithReference(styleId?.Id, state);
    }
}
