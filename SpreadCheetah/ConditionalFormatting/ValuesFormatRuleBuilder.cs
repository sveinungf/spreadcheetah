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
    /// Creates a rule that applies to cells with a value between the specified values.
    /// </summary>
    public NumericRangeFormatRuleBuilder Between(double firstValue, double secondValue) => new(firstValue, secondValue, isBetween: true);

    /// <summary>
    /// Creates a rule that applies to cells with a value not between the specified values.
    /// </summary>
    public NumericRangeFormatRuleBuilder NotBetween(double firstValue, double secondValue) => new(firstValue, secondValue, isBetween: false);

    /// <summary>
    /// Creates a rule that applies to cells with a value greater than the specified value.
    /// </summary>
    public NumericComparisonFormatRuleBuilder GreaterThan(double value) => Create(value, NumericComparisonOperator.GreaterThan);

    /// <summary>
    /// Creates a rule that applies to cells with a value less than the specified value.
    /// </summary>
    public NumericComparisonFormatRuleBuilder LessThan(double value) => Create(value, NumericComparisonOperator.LessThan);

    /// <summary>
    /// Creates a rule that applies to cells with a value equal to the specified value.
    /// </summary>
    public NumericComparisonFormatRuleBuilder EqualTo(double value) => Create(value, NumericComparisonOperator.EqualTo);

    /// <summary>
    /// Creates a rule that applies to cells with a value not equal to the specified value.
    /// </summary>
    public NumericComparisonFormatRuleBuilder NotEqualTo(double value) => Create(value, NumericComparisonOperator.NotEqualTo);

    /// <summary>
    /// Creates a rule that applies to cells with a value greater than or equal to the specified value.
    /// </summary>
    public NumericComparisonFormatRuleBuilder GreaterThanOrEqualTo(double value) => Create(value, NumericComparisonOperator.GreaterThanOrEqualTo);

    /// <summary>
    /// Creates a rule that applies to cells with a value less than or equal to the specified value.
    /// </summary>
    public NumericComparisonFormatRuleBuilder LessThanOrEqualTo(double value) => Create(value, NumericComparisonOperator.LessThanOrEqualTo);

    private static NumericComparisonFormatRuleBuilder Create(double value, NumericComparisonOperator op)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
            ThrowHelper.ValueNotFinite(nameof(value));

        return new(value, op);
    }
}
