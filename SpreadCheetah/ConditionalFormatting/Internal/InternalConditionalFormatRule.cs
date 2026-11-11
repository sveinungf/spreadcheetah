using SpreadCheetah.CellReferences;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal closed class InternalConditionalFormatRule
{
    public required int? StyleDxfId { get; init; }

    public abstract bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell);
}
