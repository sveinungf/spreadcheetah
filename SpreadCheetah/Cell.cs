using System.Globalization;
using System.Net;

namespace SpreadCheetah
{
    public class Cell
    {
        public CellDataType DataType { get; }
        public string Value { get; }

        public Cell(string? value)
        {
            DataType = CellDataType.InlineString;
            Value = value != null ? WebUtility.HtmlEncode(value) : string.Empty;
        }

        public Cell(int value)
        {
            DataType = CellDataType.Number;
            Value = value.ToString(CultureInfo.InvariantCulture);
        }

        public Cell(int? value)
        {
            DataType = CellDataType.Number;
            Value = value?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        }

        public Cell(float value)
        {
            DataType = CellDataType.Number;
            Value = value.ToString("G7", CultureInfo.InvariantCulture);
        }

        public Cell(float? value)
        {
            DataType = CellDataType.Number;
            Value = value?.ToString("G7", CultureInfo.InvariantCulture) ?? string.Empty;
        }

        public Cell(double value)
        {
            DataType = CellDataType.Number;
            Value = value.ToString("G15", CultureInfo.InvariantCulture);
        }

        public Cell(double? value)
        {
            DataType = CellDataType.Number;
            Value = value?.ToString("G15", CultureInfo.InvariantCulture) ?? string.Empty;
        }

        public Cell(decimal value)
            : this(decimal.ToDouble(value))
        {
        }

        public Cell(decimal? value)
            : this(value != null ? decimal.ToDouble(value.Value) : null as double?)
        {
        }

        public Cell(bool value)
        {
            DataType = CellDataType.Boolean;
            Value = GetStringValue(value);
        }

        public Cell(bool? value)
        {
            DataType = CellDataType.Boolean;
            Value = value != null ? GetStringValue(value.Value) : string.Empty;
        }

        private static string GetStringValue(bool value) => value ? "1" : "0";
    }
}
