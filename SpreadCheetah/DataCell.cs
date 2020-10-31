using System;
using System.Globalization;
using System.Net;

namespace SpreadCheetah
{
    public readonly struct DataCell : IEquatable<DataCell>
    {
        public CellDataType DataType { get; }
        public string Value { get; }

        public DataCell(string? value)
        {
            DataType = CellDataType.InlineString;
            Value = value != null ? WebUtility.HtmlEncode(value) : string.Empty;
        }

        public DataCell(int value)
        {
            DataType = CellDataType.Number;
            Value = value.ToString(CultureInfo.InvariantCulture);
        }

        public DataCell(int? value)
        {
            DataType = CellDataType.Number;
            Value = value?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        }

        public DataCell(float value)
        {
            DataType = CellDataType.Number;
            Value = value.ToString("G7", CultureInfo.InvariantCulture);
        }

        public DataCell(float? value)
        {
            DataType = CellDataType.Number;
            Value = value?.ToString("G7", CultureInfo.InvariantCulture) ?? string.Empty;
        }

        public DataCell(double value)
        {
            DataType = CellDataType.Number;
            Value = value.ToString("G15", CultureInfo.InvariantCulture);
        }

        public DataCell(double? value)
        {
            DataType = CellDataType.Number;
            Value = value?.ToString("G15", CultureInfo.InvariantCulture) ?? string.Empty;
        }

        public DataCell(decimal value)
            : this(decimal.ToDouble(value))
        {
        }

        public DataCell(decimal? value)
            : this(value != null ? decimal.ToDouble(value.Value) : null as double?)
        {
        }

        public DataCell(bool value)
        {
            DataType = CellDataType.Boolean;
            Value = GetStringValue(value);
        }

        public DataCell(bool? value)
        {
            DataType = CellDataType.Boolean;
            Value = value != null ? GetStringValue(value.Value) : string.Empty;
        }

        private static string GetStringValue(bool value) => value ? "1" : "0";

        public override bool Equals(object? obj) => obj is DataCell cell && Equals(cell);
        public bool Equals(DataCell other) => DataType == other.DataType && string.Equals(Value, other.Value, StringComparison.Ordinal);

        public override int GetHashCode()
        {
            var hashCode = -1382053921;
            hashCode = hashCode * -1521134295 + DataType.GetHashCode();
            hashCode = hashCode * -1521134295 + StringComparer.Ordinal.GetHashCode(Value);
            return hashCode;
        }

        public static bool operator ==(DataCell left, DataCell right) => left.Equals(right);
        public static bool operator !=(DataCell left, DataCell right) => !left.Equals(right);
    }
}
