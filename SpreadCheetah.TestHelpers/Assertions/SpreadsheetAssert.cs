using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SpreadCheetah.TestHelpers.Collections;
using Xunit;

namespace SpreadCheetah.TestHelpers.Assertions;

public static class SpreadsheetAssert
{
    public static ISpreadsheetAssertSheet SingleSheet(Stream stream)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        Valid(stream);

#pragma warning disable CA2000 // Dispose objects before losing scope
        var workbook = new XLWorkbook(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope
        var sheet = workbook.Worksheets.Single();
        return new ClosedXmlAssertSheet(workbook, sheet);
    }

    public static IWorksheetList Sheets(Stream stream)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        Valid(stream);

#pragma warning disable CA2000 // Dispose objects before losing scope
        var workbook = new XLWorkbook(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope
        return new ClosedXmlWorksheetList(workbook);
    }

    public static void Valid(Stream stream)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        stream.Position = 0;

        using var document = SpreadsheetDocument.Open(stream, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(document);
        Assert.Empty(errors);

        stream.Position = 0;
    }

    public static ISpreadsheetAssertDocumentProperties DocumentProperties(Stream stream)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        Valid(stream);

#pragma warning disable CA2000 // Dispose objects before losing scope
        var workbook = new XLWorkbook(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope
        return new ClosedXmlAssertDocumentProperties(workbook.Properties);
    }
}
