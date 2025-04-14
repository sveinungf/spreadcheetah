using ClosedXML.Excel;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class ClosedXmlAssertDocumentProperties(XLWorkbookProperties properties)
    : ISpreadsheetAssertDocumentProperties
{
    public string? Author => properties.Author;
    public string? Subject => properties.Subject;
    public string? Title => properties.Title;
    public DateTime? Created => properties.Created == default ? null : properties.Created.ToUniversalTime();
}
