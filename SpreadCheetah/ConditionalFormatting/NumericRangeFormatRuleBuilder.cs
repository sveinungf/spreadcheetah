namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Builder for a numeric range conditional format rule.
/// Use <see cref="ConditionalFormatRule.Values"/> to create an instance of this builder.
/// </summary>
public sealed class NumericRangeFormatRuleBuilder
{
    internal double FirstValue { get; }
    internal double SecondValue { get; }
    internal bool IsBetween { get; }

    internal NumericRangeFormatRuleBuilder(double firstValue, double secondValue, bool isBetween)
    {
        FirstValue = firstValue;
        SecondValue = secondValue;
        IsBetween = isBetween;
    }

    /// <summary>
    /// Sets the style that should be applied to cells matching the rule.
    /// </summary>
    /// <returns>The conditional format rule.</returns>
    public NumericRangeFormatRule WithStyle(ConditionalFormatStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new NumericRangeFormatRule(FirstValue, SecondValue, IsBetween) { Style = style };
    }
}
