using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using SpreadCheetah.TestHelpers.Assertions;
using System.Collections;

namespace SpreadCheetah.TestHelpers.Collections;

internal sealed class ClosedXmlWorksheetList(XLWorkbook workbook, SpreadsheetDocument document) : IWorksheetList
{
    private readonly List<ISpreadsheetAssertSheet> _sheets = CreateSheets(workbook, document);

    private static List<ISpreadsheetAssertSheet> CreateSheets(XLWorkbook workbook, SpreadsheetDocument document)
    {
        if (document.WorkbookPart is null)
            throw new InvalidOperationException("The document does not contain a WorkbookPart.");

        var openXmlWorksheets = document.WorkbookPart.WorksheetParts
            .Select(wp => wp.Worksheet)
            .ToList();

        return
        [
            ..workbook.Worksheets
                .Index()
                .Select(x => new ClosedXmlAssertSheet(workbook, x.Item, openXmlWorksheets[x.Index]))
        ];
    }

    public ISpreadsheetAssertSheet this[int index] => _sheets[index];
    public int Count => _sheets.Count;
    public IEnumerator<ISpreadsheetAssertSheet> GetEnumerator() => _sheets.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => _sheets.GetEnumerator();

    public void Dispose()
    {
        workbook.Dispose();
        document.Dispose();
    }
}