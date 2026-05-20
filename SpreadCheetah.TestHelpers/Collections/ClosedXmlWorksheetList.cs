using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Polyfills;
using SpreadCheetah.TestHelpers.Implementations;
using SpreadCheetah.TestHelpers.Interfaces;
using System.Collections;

namespace SpreadCheetah.TestHelpers.Collections;

internal sealed class ClosedXmlWorksheetList(XLWorkbook workbook, SpreadsheetDocument document) : IWorksheetList
{
    private readonly List<IWorksheet> _sheets = CreateSheets(workbook, document);

    private static List<IWorksheet> CreateSheets(XLWorkbook workbook, SpreadsheetDocument document)
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
                .Select(x => new ClosedXmlWorksheet(workbook, x.Item, openXmlWorksheets[x.Index]))
        ];
    }

    public IWorksheet this[int index] => _sheets[index];
    public int Count => _sheets.Count;
    public IEnumerator<IWorksheet> GetEnumerator() => _sheets.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => _sheets.GetEnumerator();

    public void Dispose()
    {
        workbook.Dispose();
        document.Dispose();
    }
}