using SpreadCheetah.CellWriters;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Number;

internal abstract class NumberCellValueWriter : NumberCellValueWriterBase
{
    public override bool WriteFormulaStartElement(StyleId? styleId, CellWriterState state)
    {
        return WriteFormulaStartElement(styleId?.Id, state.Buffer);
    }

    public override bool WriteFormulaStartElementWithReference(StyleId? styleId, CellWriterState state)
    {
        return WriteFormulaStartElementWithReference(styleId?.Id, state);
    }
}
