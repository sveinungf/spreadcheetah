using System;
using System.Globalization;
using System.Net;

namespace SpreadCheetah
{
    public readonly struct DataCell : IEquatable<DataCell>
    {
        /// <summary>
        /// The cell's Open XML data type.
        /// </summary>
        public CellDataType DataType { get; }

        /// <summary>
        /// The XML encoded value for the cell. Note that the value can differ from the original value
        /// passed to the constructor due to XML encoding or change in number precision.
        /// </summary>
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

        public DataCell(long value) : this((double)value)
        {
        }

        public DataCell(long? value) : this((double?)value)
        {
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

        public DataCell(decimal value) : this(decimal.ToDouble(value))
        {
        }

        public DataCell(decimal? value) : this(value != null ? decimal.ToDouble(value.Value) : null as double?)
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

        /// <inheritdoc/>
        public bool Equals(DataCell other) => DataType == other.DataType && string.Equals(Value, other.Value, StringComparison.Ordinal);

        /// <inheritdoc/>
        public override bool Equals(object? obj) => obj is DataCell cell && Equals(cell);

        /// <inheritdoc/>
        public override int GetHashCode() => HashCode.Combine(DataType, Value);

        /// <summary>Determines whether two instances have the same value.</summary>
        public static bool operator ==(DataCell left, DataCell right) => left.Equals(right);

        /// <summary>Determines whether two instances have different values.</summary>
        public static bool operator !=(DataCell left, DataCell right) => !left.Equals(right);
    }
}
