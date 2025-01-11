using SpreadCheetah.Tables;

namespace SpreadCheetah.TestHelpers.Assertions;

public interface ISpreadsheetAssertTableColumn
{
    string Name { get; }
    string? TotalRowLabel { get; }
    TableTotalRowFunction? TotalRowFunction { get; }
}