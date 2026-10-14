using SpreadCheetah.ConditionalFormatting;
using SpreadCheetah.TestHelpers.Models;

namespace SpreadCheetah.Test.Extensions;

internal static class ValuesFormatRuleBuilderExtensions
{
    extension(ValuesFormatRuleBuilder builder)
    {
        public NumericComparisonFormatRuleBuilder For(ConditionalFormatOperator op, double value)
        {
            return op switch
            {
                ConditionalFormatOperator.GreaterThan => builder.GreaterThan(value),
                ConditionalFormatOperator.LessThan => builder.LessThan(value),
                ConditionalFormatOperator.EqualTo => builder.EqualTo(value),
                ConditionalFormatOperator.NotEqualTo => builder.NotEqualTo(value),
                ConditionalFormatOperator.GreaterThanOrEqualTo => builder.GreaterThanOrEqualTo(value),
                ConditionalFormatOperator.LessThanOrEqualTo => builder.LessThanOrEqualTo(value),
                _ => throw new ArgumentOutOfRangeException(nameof(op), op, null)
            };
        }
    }
}
