using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents a conditional format rule that applies to cells based on a comparison to the average value.
/// Use the builder from <see cref="ConditionalFormatRule.Values"/> to create an instance of this class.
/// </summary>
public sealed class AverageComparisonFormatRule : ConditionalFormatRule
{
    internal AverageComparisonOperator Operator { get; }

    internal AverageComparisonFormatRule(AverageComparisonOperator op)
    {
        Operator = op;
    }

    internal override InternalConditionalFormatRule ToInternal(int? styleDxfId)
    {
        return new InternalAverageComparisonFormatRule
        {
            StyleDxfId = styleDxfId,
            Operator = Operator
        };
    }
}
