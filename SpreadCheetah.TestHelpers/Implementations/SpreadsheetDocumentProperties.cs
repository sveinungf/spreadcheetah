using ClosedXML.Excel;
using DocumentFormat.OpenXml.ExtendedProperties;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Implementations;

internal sealed class SpreadsheetDocumentProperties(
    XLWorkbookProperties closedXmlProperties,
    Properties? openXmlExtendedProperties)
    : ISpreadsheetDocumentProperties
{
    public string? Application => openXmlExtendedProperties?.Application?.InnerText;
    public string? Author => closedXmlProperties.Author;
    public string? Subject => closedXmlProperties.Subject;
    public string? Title => closedXmlProperties.Title;
    public DateTime? Created => closedXmlProperties.Created == default ? null : closedXmlProperties.Created.ToUniversalTime();
}
