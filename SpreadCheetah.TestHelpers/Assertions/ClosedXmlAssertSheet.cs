using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertSheet(XLWorkbook workbook, IXLWorksheet sheet)
    : ISpreadsheetAssertSheet
{
    public ISpreadsheetAssertCell this[string reference]
    {
        get
        {
            if (!SpreadsheetUtility.TryParseColumnName(reference.AsSpan(0, 1), out _))
                throw new ArgumentException("Invalid reference.", nameof(reference));
            if (!int.TryParse(reference.Substring(1), out var rowNumber))
                throw new ArgumentException("Invalid reference. The column name must currently have a single letter.", nameof(reference));

            return this[reference.Substring(0, 1), rowNumber];
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
