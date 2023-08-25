using ClosedXML.Excel;

namespace SpreadCheetah.Test.Helpers;

internal static class ClosedXmlExtensions
{
    public static string? GetCellNoteText(this IXLCell cell) => cell.HasComment ? cell.GetComment().Text : null;
}
