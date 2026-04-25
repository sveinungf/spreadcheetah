using SpreadCheetah.CellReferences;
using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

public abstract class ConditionalFormatRule
{
    public static UniqueValuesFormatRuleBuilder UniqueValues() => new();

    internal abstract ImmutableConditionalFormatRule ToImmutable(SingleCellOrCellRangeReference reference);

    private protected ConditionalFormatRule()
    {
    }
}
