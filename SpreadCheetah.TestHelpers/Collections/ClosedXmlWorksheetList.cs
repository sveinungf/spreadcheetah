using ClosedXML.Excel;
using SpreadCheetah.TestHelpers.Assertions;
using System.Collections;

namespace SpreadCheetah.TestHelpers.Collections;

internal sealed class ClosedXmlWorksheetList(XLWorkbook workbook) : IWorksheetList
{
    private readonly List<ISpreadsheetAssertSheet> _list =
    [
        .. workbook.Worksheets.Select(x => new ClosedXmlAssertSheet(workbook, x))
    ];

    public ISpreadsheetAssertSheet this[int index] => _list[index];
    public int Count => _list.Count;
    public IEnumerator<ISpreadsheetAssertSheet> GetEnumerator() => _list.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => _list.GetEnumerator();

    public void Dispose() => workbook.Dispose();
}