using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters;

internal sealed class NullValueWriter : NullValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.Id;

    // <c/>
    protected override ReadOnlySpan<byte> NullDataCell() => new[]
    {
        (byte)'<', (byte)'c', (byte)'/', (byte)'>'
    };

    protected override ReadOnlySpan<byte> BeginFormulaCell() => FormulaCellHelper.BeginNumberFormulaCell;
}
