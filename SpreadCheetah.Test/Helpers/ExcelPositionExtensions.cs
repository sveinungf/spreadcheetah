using OfficeOpenXml.Drawing;

namespace SpreadCheetah.Test.Helpers;

internal static class ExcelPositionExtensions
{
    public static string ToCellReferenceString(this ExcelDrawing.ExcelPosition position)
    {
        return $"{SpreadsheetUtility.GetColumnName(position.Column + 1)}{position.Row + 1}";
    }
}
