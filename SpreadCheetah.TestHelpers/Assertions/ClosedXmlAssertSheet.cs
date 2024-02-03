using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertSheet(XLWorkbook workbook, IXLWorksheet sheet)
    : ISpreadsheetAssertSheet
{
    public IEnumerable<ISpreadsheetAssertCell> this[string columnName]
    {
        get
        {
            var cells = sheet.Column(columnName).CellsUsed();
            return cells.Select(x => new ClosedXmlAssertCell(x));
        }
    }

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

    public void Dispose() => workbook.Dispose();
}
