using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;

namespace SpreadCheetah.CellValueWriters.Number;

internal abstract class NumberCellValueWriter : NumberCellValueWriterBase
{
    protected override int GetStyleId(StyleId styleId) => styleId.Id;

    // <c><v>
    protected override ReadOnlySpan<byte> BeginDataCell() => new[]
    {
        (byte)'<', (byte)'c', (byte)'>', (byte)'<', (byte)'v', (byte)'>'
    };

    protected override ReadOnlySpan<byte> BeginFormulaCell() => FormulaCellHelper.BeginNumberFormulaCell;
}
