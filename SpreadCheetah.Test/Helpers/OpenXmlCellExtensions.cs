using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadCheetah.Test.Helpers
{
    internal static class OpenXmlCellExtensions
    {
        public static CellValues GetDataType(this DocumentFormat.OpenXml.Spreadsheet.Cell cell)
        {
            // The "t" attribute on cells is optional and defaults to "n" (number)
            return cell.DataType?.Value ?? CellValues.Number;
        }
    }
}
