using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents a conditional format rule that applies to cells with unique values.
/// Use the builder from <see cref="ConditionalFormatRule.UniqueValues"/> to create an instance of this class.
/// </summary>
public sealed class UniqueValuesFormatRule : ConditionalFormatRule
{
    internal override InternalConditionalFormatRule ToInternal(int? styleDxfId)
    {
        return new InternalUniqueValuesFormatRule
        {
            StyleDxfId = styleDxfId
        };
    }
}
