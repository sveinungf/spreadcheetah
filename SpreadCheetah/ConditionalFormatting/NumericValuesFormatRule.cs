using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents a conditional format rule that applies to cells based on a comparison of their numeric value.
/// Use the builder from <see cref="ConditionalFormatRule.Values"/> to create an instance of this class.
/// </summary>
public sealed class NumericValuesFormatRule : ConditionalFormatRule
{
    internal double Value { get; }
    internal NumericValueOperator Operator { get; }

    internal NumericValuesFormatRule(double value, NumericValueOperator op)
    {
        Value = value;
        Operator = op;
    }

    internal override InternalConditionalFormatRule ToInternal(int? styleDxfId)
    {
        return new InternalNumericValuesFormatRule
        {
            StyleDxfId = styleDxfId,
            Value = Value,
            Operator = Operator
        };
    }
}
