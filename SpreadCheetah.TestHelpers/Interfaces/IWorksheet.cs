namespace SpreadCheetah.TestHelpers.Interfaces;

public interface IWorksheet : IDisposable
{
    string Name { get; }

    IWorksheetCell this[string reference] { get; }
    IWorksheetCell this[string columnName, int rowNumber] { get; }
    int CellCount { get; }
    int RowCount { get; }
    int? MaxRowOutlineLevel { get; }

    IWorksheetColumn Column(string columnName);
    IReadOnlyList<IWorksheetColumn> Columns { get; }
    IWorksheetRow Row(int rowNumber);

    IReadOnlyList<IWorksheetTable> Tables { get; }
}
