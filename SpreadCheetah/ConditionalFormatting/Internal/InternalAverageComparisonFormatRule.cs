using SpreadCheetah.CellReferences;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed class InternalAverageComparisonFormatRule : InternalConditionalFormatRule
{
    internal AverageComparisonOperator Operator { get; set; }

    public override bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell)
    {
        throw new NotImplementedException();
    }
}
