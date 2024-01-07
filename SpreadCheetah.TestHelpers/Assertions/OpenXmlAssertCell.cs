using System.Globalization;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace SpreadCheetah.TestHelpers.Assertions;

internal sealed class OpenXmlAssertCell(OpenXmlCell cell) : ISpreadsheetAssertCell
{
    public int? IntValue
    {
        get
        {
            if (cell.CellValue is not { } value)
                throw new InvalidOperationException($"{nameof(cell.CellValue)} was null");

            if (value.Text is null)
                return null;

            if (!int.TryParse(value.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var intValue))
                throw new InvalidOperationException($"The value {value.Text} could not be parsed as an integer");

            return intValue;
        }
    }

    public decimal? DecimalValue
    {
        get
        {
            if (cell.CellValue is not { } value)
                throw new InvalidOperationException($"{nameof(cell.CellValue)} was null");

            if (value.Text is null)
                return null;

            if (!decimal.TryParse(value.Text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var decimalValue))
                throw new InvalidOperationException($"The value {value.Text} could not be parsed as a decimal");

            return decimalValue;
        }
    }

    public string? StringValue
    {
        get
        {
            if (cell.InlineString is not { } inlineString)
                throw new InvalidOperationException($"{nameof(cell.InlineString)} was null");

            if (inlineString.Text is not { } text)
                throw new InvalidOperationException($"{nameof(inlineString.Text)} was null");

            return text.Text;
        }
    }
}
