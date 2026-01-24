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

        using var openXmlDoc = GetValidatedOpenXmlDocument(stream);
        var openXmlWorksheet = openXmlDoc.WorkbookPart?.WorksheetParts.Single().Worksheet
            ?? throw new InvalidOperationException("No worksheet found in the document.");

#pragma warning disable CA2000 // Dispose objects before losing scope
        var workbook = new XLWorkbook(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope
        var sheet = workbook.Worksheets.Single();
        return new ClosedXmlAssertSheet(workbook, sheet, openXmlWorksheet);
    }

    public static IWorksheetList Sheets(Stream stream)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

#pragma warning disable CA2000 // Dispose objects before losing scope
        var openXmlDoc = GetValidatedOpenXmlDocument(stream);
        var workbook = new XLWorkbook(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope

        return new ClosedXmlWorksheetList(workbook, openXmlDoc);
    }

    public static void Valid(Stream stream)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        using var _ = GetValidatedOpenXmlDocument(stream);
    }

    public static ISpreadsheetAssertDocumentProperties DocumentProperties(Stream stream)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

#pragma warning disable CA2000 // Dispose objects before losing scope
        var closedXmlWorkbook = new XLWorkbook(stream);
        var openXmlDoc = GetValidatedOpenXmlDocument(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope

        return new AssertDocumentProperties(
            closedXmlWorkbook.Properties,
            openXmlDoc.ExtendedFilePropertiesPart?.Properties);
    }

    private static SpreadsheetDocument GetValidatedOpenXmlDocument(Stream stream)
    {
        stream.Position = 0;
        var openXmlDoc = SpreadsheetDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(openXmlDoc);
        Assert.Empty(errors);

        stream.Position = 0;
        return openXmlDoc;
    }
}
