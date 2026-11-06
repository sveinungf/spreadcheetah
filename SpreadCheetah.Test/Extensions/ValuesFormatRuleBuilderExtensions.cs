using SpreadCheetah.ConditionalFormatting;
using SpreadCheetah.TestHelpers.Models;

namespace SpreadCheetah.Test.Extensions;

internal static class ValuesFormatRuleBuilderExtensions
{
    extension(ValuesFormatRuleBuilder builder)
    {
        public AverageComparisonFormatRuleBuilder For(AverageComparisonOperator op)
        {
            return op switch
            {
                AverageComparisonOperator.Above => builder.AboveAverage(),
                AverageComparisonOperator.Below => builder.BelowAverage(),
                AverageComparisonOperator.EqualToOrAbove => builder.EqualToOrAboveAverage(),
                AverageComparisonOperator.EqualToOrBelow => builder.EqualToOrBelowAverage(),
                AverageComparisonOperator.OneStdDevAbove => builder.StandardDeviationsAboveAverage(1),
                AverageComparisonOperator.OneStdDevBelow => builder.StandardDeviationsBelowAverage(1),
                AverageComparisonOperator.TwoStdDevAbove => builder.StandardDeviationsAboveAverage(2),
                AverageComparisonOperator.TwoStdDevBelow => builder.StandardDeviationsBelowAverage(2),
                AverageComparisonOperator.ThreeStdDevAbove => builder.StandardDeviationsAboveAverage(3),
                AverageComparisonOperator.ThreeStdDevBelow => builder.StandardDeviationsBelowAverage(3),
                _ => throw new ArgumentOutOfRangeException(nameof(op), op, null)
            };
        }

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
