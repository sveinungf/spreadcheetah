using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellValueWriters.Time;

internal sealed class NullDateTimeCellValueWriter : NullValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.DateTimeId;

    public override bool TryWriteCell(in DataCell cell, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        var defaultStyleId = defaultStyling?.DateTimeStyleId;
        return defaultStyleId is not null
            ? TryWriteCell(defaultStyleId.Value, buffer)
            : TryWriteCell(buffer);
    }
}
