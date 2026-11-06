using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.TestHelpers.Models;

namespace SpreadCheetah.TestHelpers.Extensions;

internal static class OpenXmlConditionalFormattingRuleExtensions
{
    extension(ConditionalFormattingRule rule)
    {
        public bool TryGetAverageComparisonOperator(out AverageComparisonOperator op)
        {
            if (rule.Type?.Value != ConditionalFormatValues.AboveAverage)
            {
                op = (AverageComparisonOperator)(-1);
                return false;
            }

            var aboveAverage = rule.AboveAverage?.Value ?? true;
            var equalAverage = rule.EqualAverage?.Value ?? false;
            var standardDeviations = rule.StdDev?.Value;

            if (standardDeviations is not null)
            {
                op = (aboveAverage, standardDeviations) switch
                {
                    (true, 1) => AverageComparisonOperator.OneStdDevAbove,
                    (false, 1) => AverageComparisonOperator.OneStdDevBelow,
                    (true, 2) => AverageComparisonOperator.TwoStdDevAbove,
                    (false, 2) => AverageComparisonOperator.TwoStdDevBelow,
                    (true, 3) => AverageComparisonOperator.ThreeStdDevAbove,
                    (false, 3) => AverageComparisonOperator.ThreeStdDevBelow,
                    _ => throw new InvalidOperationException($"Unexpected average comparison rule: aboveAverage={aboveAverage}, stdDev={standardDeviations}.")
                };

                return true;
            }

            op = (aboveAverage, equalAverage) switch
            {
                (true, false) => AverageComparisonOperator.Above,
                (false, false) => AverageComparisonOperator.Below,
                (true, true) => AverageComparisonOperator.EqualToOrAbove,
                _ => AverageComparisonOperator.EqualToOrBelow
            };

            return true;
        }
    }
}
