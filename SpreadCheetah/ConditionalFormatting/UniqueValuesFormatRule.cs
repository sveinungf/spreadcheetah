using SpreadCheetah.CellReferences;
using SpreadCheetah.ConditionalFormatting.Internal;
using SpreadCheetah.Styling;

namespace SpreadCheetah.ConditionalFormatting;

public sealed class UniqueValuesFormatRule : ConditionalFormatRule
{
    private readonly Style _style;

    internal UniqueValuesFormatRule(Style style)
    {
        _style = style;
    }

    internal override ImmutableConditionalFormatRule ToImmutable(SingleCellOrCellRangeReference reference)
    {
        return new ImmutableUniqueValuesFormatRule
        {
            Reference = reference,
            Style = _style
        };
    }
}
