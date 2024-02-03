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

    public IEnumerable<ISpreadsheetAssertCell> this[string columnName]
    {
        get
        {
            var cells = sheet.Column(columnName).CellsUsed();
            return cells.Select(x => new ClosedXmlAssertCell(x));
        }
    }

    public int CellCount => sheet.CellsUsed().Count();

    public int RowCount => sheet.RowsUsed().Count();

    public IEnumerable<ISpreadsheetAssertCell> Row(int rowNumber)
    {
        var cells = sheet.Row(rowNumber).CellsUsed();
        return cells.Select(x => new ClosedXmlAssertCell(x));
    }

    public void Dispose() => workbook.Dispose();
}
