namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IConditionalFormatRule
{
    string CellRangeReference { get; }
    bool IsDuplicateValuesRule { get; }
    bool IsUniqueValuesRule { get; }
    IStyle Style { get; }
}
