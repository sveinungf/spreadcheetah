namespace SpreadCheetah.TestHelpers.Assertions;

public interface IConditionalFormatRule
{
    string CellRangeReference { get; }
    bool IsUniqueValuesRule { get; }
    ISpreadsheetAssertStyle Style { get; }
}
