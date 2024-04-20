using OpenXmlCellValue = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace SpreadCheetah.Test.Helpers;

internal static class OpenXmlCellExtensions
{
    public static OpenXmlCellValue GetDataType(this DocumentFormat.OpenXml.Spreadsheet.Cell cell)
    {
        // The "t" attribute on cells is optional and defaults to "n" (number)
        return cell.DataType?.Value ?? OpenXmlCellValue.Number;
    }
}
