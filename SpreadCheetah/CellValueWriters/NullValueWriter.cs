using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters;

internal sealed class NullValueWriter : NullValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.Id;

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(buffer);
    }

    public override bool TryWriteCell(string formulaText, in DataCell cachedValue, StyleId? styleId, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        return TryWriteCell(formulaText, styleId?.Id, buffer);
    }
}
