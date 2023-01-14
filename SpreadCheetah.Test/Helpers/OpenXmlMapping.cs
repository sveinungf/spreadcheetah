using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.Test.Helpers;

internal static class OpenXmlMapping
{
    public static SheetStateValues GetOpenXmlWorksheetVisibility(this WorksheetVisibility visibility)
    {
        return visibility switch
        {
            WorksheetVisibility.Visible => SheetStateValues.Visible,
            WorksheetVisibility.Hidden => SheetStateValues.Hidden,
            WorksheetVisibility.VeryHidden => SheetStateValues.VeryHidden,
            _ => throw new NotImplementedException(),
        };
    }
}
