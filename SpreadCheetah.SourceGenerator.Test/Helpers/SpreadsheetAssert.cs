using DocumentFormat.OpenXml.Packaging;
using SpreadCheetah.SourceGenerator.Test.Helpers.OpenXml;

namespace SpreadCheetah.SourceGenerator.Test.Helpers;

internal static class SpreadsheetAssert
{
    public static ISpreadsheetAssertSheet SingleSheet(Stream stream)
    {
        stream.Position = 0;
#pragma warning disable CA2000 // Dispose objects before losing scope
        var document = SpreadsheetDocument.Open(stream, false);
#pragma warning restore CA2000 // Dispose objects before losing scope
        var sheetPart = document.WorkbookPart!.WorksheetParts.Single();
        return new OpenXmlAssertSheet(document, sheetPart.Worksheet);
    }
}
