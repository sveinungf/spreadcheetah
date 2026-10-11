using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

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
