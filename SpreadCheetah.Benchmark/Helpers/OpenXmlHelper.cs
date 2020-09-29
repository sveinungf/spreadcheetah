using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace SpreadCheetah.Benchmark.Helpers
{
    internal static class OpenXmlHelper
    {
        public static OpenXmlCell CreateCell(string value)
        {
            var inlineString = new InlineString();
            inlineString.AppendChild(new Text(value));
            var cell = new OpenXmlCell { DataType = CellValues.InlineString };
            cell.AppendChild(inlineString);
            return cell;
        }
    }
}
