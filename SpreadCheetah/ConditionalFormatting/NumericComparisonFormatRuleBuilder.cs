using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Builder for a numeric values conditional format rule.
/// Use <see cref="ConditionalFormatRule.Values"/> to create an instance of this builder.
/// </summary>
public sealed class NumericComparisonFormatRuleBuilder
{
    internal double Value { get; }
    internal NumericComparisonOperator Operator { get; }

    internal NumericComparisonFormatRuleBuilder(double value, NumericComparisonOperator op)
    {
        Value = value;
        Operator = op;
    }

    /// <summary>
    /// Sets the style that should be applied to cells matching the rule.
    /// </summary>
    /// <returns>The conditional format rule.</returns>
    public NumericComparisonFormatRule WithStyle(ConditionalFormatStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new NumericComparisonFormatRule(Value, Operator) { Style = style };
    }
}
