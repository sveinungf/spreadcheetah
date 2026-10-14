using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents a conditional format rule that applies to cells based on a comparison of their numeric value.
/// Use the builder from <see cref="ConditionalFormatRule.Values"/> to create an instance of this class.
/// </summary>
public sealed class NumericComparisonFormatRule : ConditionalFormatRule
{
    internal double Value { get; }
    internal NumericComparisonOperator Operator { get; }

    internal NumericComparisonFormatRule(double value, NumericComparisonOperator op)
    {
        Value = value;
        Operator = op;
    }

    internal override InternalConditionalFormatRule ToInternal(int? styleDxfId)
    {
        return new InternalNumericComparisonFormatRule
        {
            StyleDxfId = styleDxfId,
            Value = Value,
            Operator = Operator
        };
    }
}
