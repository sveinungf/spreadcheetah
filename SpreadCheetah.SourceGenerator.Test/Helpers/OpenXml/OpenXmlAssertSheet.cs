using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace SpreadCheetah.SourceGenerator.Test.Helpers.OpenXml;

internal sealed class OpenXmlAssertSheet(SpreadsheetDocument document, Worksheet worksheet)
    : ISpreadsheetAssertSheet
{
    private List<List<OpenXmlCell>>? _cells;
    private List<List<OpenXmlCell>> Cells => _cells ??= GetAllCells();

    private List<List<OpenXmlCell>> GetAllCells()
    {
        var allCells = new List<List<OpenXmlCell>>();

        foreach (var row in worksheet.Descendants<Row>())
        {
            var rowCells = row.Descendants<OpenXmlCell>().ToList();
            allCells.Add(rowCells);
        }

        return allCells;
    }

    public ISpreadsheetAssertCell this[string columnName, int rowNumber]
    {
        get
        {
            if (!SpreadsheetUtility.TryParseColumnName(columnName.AsSpan(), out var columnNumber))
                throw new InvalidOperationException($"{columnName} is not a valid column name");

            if (Cells.ElementAtOrDefault(rowNumber - 1) is not { } row)
                throw new InvalidOperationException($"Could not find row number {rowNumber}");

            if (row.ElementAtOrDefault(columnNumber - 1) is not { } cell)
                throw new InvalidOperationException($"Could not find cell with reference {columnName}{rowNumber}");

            return new OpenXmlAssertCell(cell);
        }
    }

    public IEnumerable<ISpreadsheetAssertCell> this[string columnName]
    {
        get
        {
            if (!SpreadsheetUtility.TryParseColumnName(columnName.AsSpan(), out var columnNumber))
                throw new InvalidOperationException($"{columnName} is not a valid column name");

            foreach (var (row, index) in Cells.Select((x, i) => (x, i)))
            {
                var rowNumber = index + 1;
                if (row.ElementAtOrDefault(columnNumber - 1) is not { } cell)
                    throw new InvalidOperationException($"Could not find cell with reference {columnName}{rowNumber}");

                yield return new OpenXmlAssertCell(cell);
            }
        }
    }

    public int CellCount => Cells.Sum(x => x.Count);
    public int RowCount => Cells.Count;

    public void Dispose()
    {
        document.Dispose();
    }
}
