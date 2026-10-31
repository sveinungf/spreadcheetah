using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Builder for an average comparison conditional format rule.
/// Use <see cref="ConditionalFormatRule.Values"/> to create an instance of this builder.
/// </summary>
public sealed class AverageComparisonFormatRuleBuilder
{
    internal AverageComparisonOperator Operator { get; }

    internal AverageComparisonFormatRuleBuilder(AverageComparisonOperator op)
    {
        Operator = op;
    }

    /// <summary>
    /// Sets the style that should be applied to cells matching the rule.
    /// </summary>
    /// <returns>The conditional format rule.</returns>
    public AverageComparisonFormatRule WithStyle(ConditionalFormatStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new AverageComparisonFormatRule(Operator) { Style = style };
    }
}
