using SpreadCheetah.CellReferences;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed class ConditionalFormatRulesManager
{
    public Dictionary<SingleCellOrCellRangeReference, List<ImmutableConditionalFormatRule>> Rules { get; } = [];

    public void AddRule(SingleCellOrCellRangeReference reference, ImmutableConditionalFormatRule rule)
    {
        if (!Rules.TryGetValue(reference, out var rules))
        {
            rules = [];
            Rules.Add(reference, rules);
        }

        rules.Add(rule);
    }
}
