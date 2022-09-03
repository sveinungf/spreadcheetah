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
}
