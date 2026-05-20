using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class ClosedXmlWorksheet(XLWorkbook workbook, IXLWorksheet sheet, Worksheet openXmlWorksheet)
    : IWorksheet
{
    public string Name => sheet.Name;

    public IWorksheetCell this[string reference]
    {
        get
        {
            if (!SpreadsheetUtility.TryParseColumnName(reference.AsSpan(0, 1), out _))
                throw new ArgumentException("Invalid reference.", nameof(reference));
            if (!int.TryParse(reference.AsSpan(1), out var rowNumber))
                throw new ArgumentException("Invalid reference. The column name must currently have a single letter.", nameof(reference));

            return this[reference[..1], rowNumber];
        }
    }

    public IWorksheetCell this[string columnName, int rowNumber]
    {
        get
        {
            var cell = sheet.Cell(rowNumber, columnName);
            return new ClosedXmlWorksheetCell(cell);
        }
    }

    public int CellCount => sheet.CellsUsed().Count();

    public int RowCount => sheet.RowsUsed().Count();

    public int? MaxRowOutlineLevel => openXmlWorksheet.SheetFormatProperties?.OutlineLevelRow?.Value;

    public IWorksheetColumn Column(string columnName)
    {
        return new ClosedXmlWorksheetColumn(sheet.Column(columnName));
    }

    public IReadOnlyList<IWorksheetColumn> Columns => field ??=
        [.. sheet.Columns().Select(x => new ClosedXmlWorksheetColumn(x))];

    public IWorksheetRow Row(int rowNumber)
    {
        var row = sheet.Row(rowNumber);
        return new ClosedXmlWorksheetRow(row);
    }

    public IReadOnlyList<IWorksheetTable> Tables => field ??=
        [.. sheet.Tables.Select(x => new ClosedXmlWorksheetTable(x))];

    public void Dispose() => workbook.Dispose();
}
