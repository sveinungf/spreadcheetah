using SpreadCheetah.TestHelpers.Models;

namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IAverageComparisonFormatRule : IConditionalFormatRule
{
    AverageComparisonOperator AverageComparisonOperator { get; }
}
