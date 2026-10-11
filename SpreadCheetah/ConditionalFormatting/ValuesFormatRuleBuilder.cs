using SpreadCheetah.ConditionalFormatting.Internal;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Builder for a values conditional format rule.
/// Use <see cref="ConditionalFormatRule.Values"/> to create an instance of this builder.
/// </summary>
public sealed class ValuesFormatRuleBuilder
{
    internal ValuesFormatRuleBuilder()
    {
    }

    /// <summary>
    /// Creates a rule that applies to cells with a value greater than the specified value.
    /// </summary>
    public NumericValuesFormatRuleBuilder GreaterThan(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
            ThrowHelper.ValueNotFinite(nameof(value));

        return new(value, NumericValueOperator.GreaterThan);
    }
}
