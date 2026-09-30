namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IConditionalFormatRule
{
    string CellRangeReference { get; }
    string Formula { get; }
    bool IsDuplicateValuesRule { get; }
    bool IsMatchesFormulaRule { get; }
    bool IsUniqueValuesRule { get; }
    IStyle Style { get; }
}
