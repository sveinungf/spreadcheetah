using SpreadCheetah.CellWriters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters;

internal sealed class NullValueWriter : NullValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.Id;

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return TryWriteCell(state);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return TryWriteCell(formulaText, styleId?.Id, state);
    }

    public override bool WriteFormulaStartElement(StyleId? styleId, DefaultStyling? defaultStyling, CellWriterState state)
    {
        return WriteFormulaStartElement(styleId?.Id, state);
    }
}
