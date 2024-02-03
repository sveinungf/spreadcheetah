using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

public static class SpreadsheetAssert
{
    public static ISpreadsheetAssertSheet SingleSheet(Stream stream)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        stream.Position = 0;
#pragma warning disable CA2000 // Dispose objects before losing scope
        var workbook = new XLWorkbook(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope
        var sheet = workbook.Worksheets.Single();
        return new ClosedXmlAssertSheet(workbook, sheet);
    }
}
