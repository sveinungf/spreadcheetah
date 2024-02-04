using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertSheet(XLWorkbook workbook, IXLWorksheet sheet)
    : ISpreadsheetAssertSheet
{
    public ISpreadsheetAssertCell this[string columnName, int rowNumber]
    {
        get
        {
            var cell = sheet.Cell(rowNumber, columnName);
            return new ClosedXmlAssertCell(cell);
        }
    }

    public int CellCount => sheet.CellsUsed().Count();

    public int RowCount => sheet.RowsUsed().Count();

    public IEnumerable<ISpreadsheetAssertCell> Column(string columnName)
    {
        var cells = sheet.Column(columnName).CellsUsed();
        return cells.Select(x => new ClosedXmlAssertCell(x));
    }

    public IEnumerable<ISpreadsheetAssertCell> Row(int rowNumber)
    {
        var row = sheet.Row(rowNumber);
        return row.IsEmpty()
            ? []
            : row.Cells(false).Select(x => new ClosedXmlAssertCell(x));
    }

    public void Dispose() => workbook.Dispose();
}
