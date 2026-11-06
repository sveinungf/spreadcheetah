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
    public NumericRangeFormatRuleBuilder Between(double firstValue, double secondValue) => CreateRange(firstValue, secondValue, isBetween: true);

    /// <summary>
    /// Creates a rule that applies to cells with a value not between the specified values.
    /// </summary>
    public NumericRangeFormatRuleBuilder NotBetween(double firstValue, double secondValue) => CreateRange(firstValue, secondValue, isBetween: false);

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

    /// <summary>
    /// Creates a rule that applies to cells with a value above the average of the range.
    /// </summary>
    public AverageComparisonFormatRuleBuilder AboveAverage() => _aboveAverage ??= new(AverageComparisonOperator.Above);
    private AverageComparisonFormatRuleBuilder? _aboveAverage;

    /// <summary>
    /// Creates a rule that applies to cells with a value below the average of the range.
    /// </summary>
    public AverageComparisonFormatRuleBuilder BelowAverage() => _belowAverage ??= new(AverageComparisonOperator.Below);
    private AverageComparisonFormatRuleBuilder? _belowAverage;

    /// <summary>
    /// Creates a rule that applies to cells with a value equal to or above the average of the range.
    /// </summary>
    public AverageComparisonFormatRuleBuilder EqualToOrAboveAverage() => _equalToOrAboveAverage ??= new(AverageComparisonOperator.EqualToOrAbove);
    private AverageComparisonFormatRuleBuilder? _equalToOrAboveAverage;

    /// <summary>
    /// Creates a rule that applies to cells with a value equal to or below the average of the range.
    /// </summary>
    public AverageComparisonFormatRuleBuilder EqualToOrBelowAverage() => _equalToOrBelowAverage ??= new(AverageComparisonOperator.EqualToOrBelow);
    private AverageComparisonFormatRuleBuilder? _equalToOrBelowAverage;

    /// <summary>
    /// Creates a rule that applies to cells with a value above the average of the range by the specified number of standard deviations.
    /// </summary>
    public AverageComparisonFormatRuleBuilder StandardDeviationsAboveAverage(int standardDeviations)
    {
        var op = standardDeviations switch
        {
            1 => AverageComparisonOperator.OneStdDevAbove,
            2 => AverageComparisonOperator.TwoStdDevAbove,
            3 => AverageComparisonOperator.ThreeStdDevAbove,
            _ => throw new ArgumentOutOfRangeException(nameof(standardDeviations), standardDeviations, "Standard deviations must be between 1 and 3.")
        };

        return new AverageComparisonFormatRuleBuilder(op);
    }

    /// <summary>
    /// Creates a rule that applies to cells with a value below the average of the range by the specified number of standard deviations.
    /// </summary>
    public AverageComparisonFormatRuleBuilder StandardDeviationsBelowAverage(int standardDeviations)
    {
        var op = standardDeviations switch
        {
            1 => AverageComparisonOperator.OneStdDevBelow,
            2 => AverageComparisonOperator.TwoStdDevBelow,
            3 => AverageComparisonOperator.ThreeStdDevBelow,
            _ => throw new ArgumentOutOfRangeException(nameof(standardDeviations), standardDeviations, "Standard deviations must be between 1 and 3.")
        };

        return new AverageComparisonFormatRuleBuilder(op);
    }

    private static NumericComparisonFormatRuleBuilder Create(double value, NumericComparisonOperator op)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
            ThrowHelper.ValueNotFinite(nameof(value));

        return new NumericComparisonFormatRuleBuilder(value, op);
    }

    private static NumericRangeFormatRuleBuilder CreateRange(double firstValue, double secondValue, bool isBetween)
    {
        if (double.IsNaN(firstValue) || double.IsInfinity(firstValue))
            ThrowHelper.ValueNotFinite(nameof(firstValue));

        if (double.IsNaN(secondValue) || double.IsInfinity(secondValue))
            ThrowHelper.ValueNotFinite(nameof(secondValue));

        return new NumericRangeFormatRuleBuilder(firstValue, secondValue, isBetween);
    }
}
