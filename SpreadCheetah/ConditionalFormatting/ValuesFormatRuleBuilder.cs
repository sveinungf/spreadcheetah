using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

public sealed class ValuesFormatRuleBuilder
{
    internal ValuesFormatRuleBuilder()
    {
    }

    public NumericValuesFormatRuleBuilder GreaterThan(double value)
    {
        return new(value, NumericValueOperator.GreaterThan);
    }
}
