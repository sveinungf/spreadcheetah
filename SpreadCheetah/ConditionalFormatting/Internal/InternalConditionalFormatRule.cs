using SpreadCheetah.CellReferences;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal abstract record InternalConditionalFormatRule
{
    public int? StyleDxfId { get; init; }

    public abstract bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell);
}
