using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

public sealed class UniqueValuesFormatRule : ConditionalFormatRule
{
    internal override ImmutableConditionalFormatRule ToImmutable(int? styleDxfId)
    {
        return new ImmutableUniqueValuesFormatRule
        {
            StyleDxfId = styleDxfId
        };
    }
}
