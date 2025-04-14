using ClosedXML.Excel;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class AssertDocumentProperties(
    XLWorkbookProperties closedXmlProperties,
    Properties? openXmlExtendedProperties)
    : ISpreadsheetAssertDocumentProperties
{
    public string? Application => openXmlExtendedProperties?.Application?.InnerText;
    public string? Author => closedXmlProperties.Author;
    public string? Subject => closedXmlProperties.Subject;
    public string? Title => closedXmlProperties.Title;
    public DateTime? Created => closedXmlProperties.Created == default ? null : closedXmlProperties.Created.ToUniversalTime();
}
