using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Time;

internal sealed class NullDateTimeCellValueWriter : NullValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.DateTimeId;

    // <c s="1"/>
    // NOTE: Assumes the default style for DateTime has index 1 in styles.xml.
    protected override ReadOnlySpan<byte> NullDataCell() => new[]
    {
        (byte)'<', (byte)'c', (byte)' ', (byte)'s', (byte)'=', (byte)'"',
        (byte)'1', (byte)'"', (byte)'/', (byte)'>'
    };

    protected override ReadOnlySpan<byte> BeginFormulaCell() => FormulaCellHelper.BeginDefaultDateTimeFormulaCell;
}
