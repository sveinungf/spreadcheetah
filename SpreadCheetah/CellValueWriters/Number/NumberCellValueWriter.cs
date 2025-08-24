using SpreadCheetah.CellWriters;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Number;

internal abstract class NumberCellValueWriter : NumberCellValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.Id;

    public override bool WriteFormulaStartElement(StyleId? styleId, CellWriterState state)
    {
        return WriteFormulaStartElement(styleId?.Id, state.Buffer);
    }

    public override bool WriteFormulaStartElementWithReference(StyleId? styleId, CellWriterState state)
    {
        return WriteFormulaStartElementWithReference(styleId?.Id, state);
    }
}
