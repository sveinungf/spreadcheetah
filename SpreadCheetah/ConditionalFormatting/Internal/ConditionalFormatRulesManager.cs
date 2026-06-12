using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed class ConditionalFormatRulesManager
{
    private int _ruleCount;

    public Dictionary<SingleCellOrCellRangeReference, List<ImmutableConditionalFormatRule>> Rules { get; } = [];

    public bool TryAddRule(SingleCellOrCellRangeReference reference, ImmutableConditionalFormatRule rule)
    {
        if (_ruleCount >= SpreadsheetConstants.MaxNumberOfConditionalFormatRules)
            return false;

        if (!Rules.TryGetValue(reference, out var rules))
        {
            rules = [];
            Rules.Add(reference, rules);
        }

        rules.Add(rule);
        _ruleCount++;
        return true;
    }
}
