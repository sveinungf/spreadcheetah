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
                throw new ArgumentException($"{columnName} is not a valid column name");

            if (Cells.ElementAtOrDefault(rowNumber - 1) is not { } row)
                throw new ArgumentException($"Could not find row number {rowNumber}");

            if (row.ElementAtOrDefault(columnNumber - 1) is not { } cell)
                throw new ArgumentException($"Could not find cell with reference {columnName}{rowNumber}");

            return new OpenXmlAssertCell(cell);
        }
    }

    public int CellCount => Cells.Sum(x => x.Count);

    public void Dispose()
    {
        document.Dispose();
    }
}
