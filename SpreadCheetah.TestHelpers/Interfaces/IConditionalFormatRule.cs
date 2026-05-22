namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IConditionalFormatRule
{
    string CellRangeReference { get; }
    bool IsUniqueValuesRule { get; }
    IStyle Style { get; }
}
