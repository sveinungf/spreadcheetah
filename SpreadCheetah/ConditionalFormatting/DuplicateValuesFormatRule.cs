using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents a conditional format rule that applies to cells with duplicate values.
/// Use the builder from <see cref="ConditionalFormatRule.DuplicateValues"/> to create an instance of this class.
/// </summary>
public sealed class DuplicateValuesFormatRule : ConditionalFormatRule
{
    internal override ImmutableConditionalFormatRule ToImmutable(int? styleDxfId)
    {
        return new ImmutableDuplicateValuesFormatRule
        {
            StyleDxfId = styleDxfId
        };
    }
}
