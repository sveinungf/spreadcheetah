using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Polyfills;
using SpreadCheetah.TestHelpers.Collections;
using SpreadCheetah.TestHelpers.Implementations;
using SpreadCheetah.TestHelpers.Interfaces;
using Xunit;

namespace SpreadCheetah.TestHelpers;

public static class SpreadsheetAssert
{
    public static IWorksheet SingleSheet(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

        using var openXmlDoc = GetValidatedOpenXmlDocument(stream);
        var openXmlWorksheet = openXmlDoc.WorkbookPart?.WorksheetParts.Single().Worksheet
            ?? throw new InvalidOperationException("No worksheet found in the document.");

#pragma warning disable CA2000 // Dispose objects before losing scope
        var workbook = new XLWorkbook(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope
        var sheet = workbook.Worksheets.Single();
        return new ClosedXmlWorksheet(workbook, sheet, openXmlWorksheet);
    }

    public static IWorksheetList Sheets(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

#pragma warning disable CA2000 // Dispose objects before losing scope
        var openXmlDoc = GetValidatedOpenXmlDocument(stream);
        var workbook = new XLWorkbook(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope

        return new ClosedXmlWorksheetList(workbook, openXmlDoc);
    }

    public static void Valid(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

        using var _ = GetValidatedOpenXmlDocument(stream);
    }

    public static ISpreadsheetDocumentProperties DocumentProperties(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

#pragma warning disable CA2000 // Dispose objects before losing scope
        var closedXmlWorkbook = new XLWorkbook(stream);
        var openXmlDoc = GetValidatedOpenXmlDocument(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope

        return new SpreadsheetDocumentProperties(
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
