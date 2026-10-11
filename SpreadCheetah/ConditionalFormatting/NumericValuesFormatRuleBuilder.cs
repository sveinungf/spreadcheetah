using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

public sealed class NumericValuesFormatRuleBuilder
{
    internal double Value { get; }
    internal NumericValueOperator Operator { get; }

    internal NumericValuesFormatRuleBuilder(double value, NumericValueOperator op)
    {
        Value = value;
        Operator = op;
    }

    public NumericValuesFormatRule WithStyle(ConditionalFormatStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new NumericValuesFormatRule(Value, Operator) { Style = style };
    }
}
