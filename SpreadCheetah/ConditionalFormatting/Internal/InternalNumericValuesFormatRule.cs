using SpreadCheetah.CellReferences;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed record InternalNumericValuesFormatRule : InternalConditionalFormatRule
{
    public required double Value { get; init; }
    public required NumericValueOperator Operator { get; init; }

    public override bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell)
    {
        // TODO: Implement
        throw new NotImplementedException();
    }
}
