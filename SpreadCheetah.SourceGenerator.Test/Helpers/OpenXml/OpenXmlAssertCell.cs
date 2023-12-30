using System.Globalization;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace SpreadCheetah.SourceGenerator.Test.Helpers.OpenXml;

internal sealed class OpenXmlAssertCell(OpenXmlCell cell) : ISpreadsheetAssertCell
{
    public int? IntValue
    {
        get
        {
            if (cell.CellValue is not { } value)
                throw new ArgumentException($"{nameof(cell.CellValue)} was null");

            if (value.Text is null)
                return null;

            if (!int.TryParse(value.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var intValue))
                throw new ArgumentException($"The value {value.Text} could not be parsed as an integer");

            return intValue;
        }
    }

    public string? StringValue
    {
        get
        {
            if (cell.InlineString is not { } inlineString)
                throw new ArgumentException($"{nameof(cell.InlineString)} was null");

            if (inlineString.Text is not { } text)
                throw new ArgumentException($"{nameof(inlineString.Text)} was null");

            return text.Text;
        }
    }
}
