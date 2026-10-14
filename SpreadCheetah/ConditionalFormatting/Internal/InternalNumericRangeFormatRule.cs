using SpreadCheetah.CellReferences;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed record InternalNumericRangeFormatRule : InternalConditionalFormatRule
{
    public required double FirstValue { get; init; }
    public required double SecondValue { get; init; }
    public required bool IsBetween { get; init; }

    public override bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell)
    {
        // TODO: Implement
        throw new NotImplementedException();
    }
}
