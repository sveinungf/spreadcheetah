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
    public NumericValuesFormatRuleBuilder GreaterThan(double value) => Create(value, NumericValueOperator.GreaterThan);

    /// <summary>
    /// Creates a rule that applies to cells with a value less than the specified value.
    /// </summary>
    public NumericValuesFormatRuleBuilder LessThan(double value) => Create(value, NumericValueOperator.LessThan);

    /// <summary>
    /// Creates a rule that applies to cells with a value equal to the specified value.
    /// </summary>
    public NumericValuesFormatRuleBuilder EqualTo(double value) => Create(value, NumericValueOperator.EqualTo);

    /// <summary>
    /// Creates a rule that applies to cells with a value not equal to the specified value.
    /// </summary>
    public NumericValuesFormatRuleBuilder NotEqualTo(double value) => Create(value, NumericValueOperator.NotEqualTo);

    /// <summary>
    /// Creates a rule that applies to cells with a value greater than or equal to the specified value.
    /// </summary>
    public NumericValuesFormatRuleBuilder GreaterThanOrEqualTo(double value) => Create(value, NumericValueOperator.GreaterThanOrEqualTo);

    /// <summary>
    /// Creates a rule that applies to cells with a value less than or equal to the specified value.
    /// </summary>
    public NumericValuesFormatRuleBuilder LessThanOrEqualTo(double value) => Create(value, NumericValueOperator.LessThanOrEqualTo);

    private static NumericValuesFormatRuleBuilder Create(double value, NumericValueOperator op)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
            ThrowHelper.ValueNotFinite(nameof(value));

        return new(value, op);
    }
}
