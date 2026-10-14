using SpreadCheetah.TestHelpers.Models;

namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IConditionalFormatRule
{
    string CellRangeReference { get; }
    string Formula { get; }
    double NumericValue { get; }
    IReadOnlyList<double> NumericValues { get; }
    ConditionalFormatOperator RuleOperator { get; }
    bool IsCellIsRule { get; }
    bool IsDuplicateValuesRule { get; }
    bool IsMatchesFormulaRule { get; }
    bool IsUniqueValuesRule { get; }
    IStyle Style { get; }
}
