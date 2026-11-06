using SpreadCheetah.Tables;

namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IWorksheetTableColumn
{
    string Name { get; }
    string? TotalRowLabel { get; }
    TableTotalRowFunction? TotalRowFunction { get; }
}