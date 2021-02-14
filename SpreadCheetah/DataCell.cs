using System;
using System.Globalization;
using System.Net;

namespace SpreadCheetah
{
    /// <summary>
    /// Represents data in a worksheet cell.
    /// </summary>
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

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a text value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// </summary>
        public DataCell(string? value)
        {
            DataType = CellDataType.InlineString;
            Value = value != null ? WebUtility.HtmlEncode(value) : string.Empty;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with an integer value.
        /// </summary>
        public DataCell(int value)
        {
            DataType = CellDataType.Number;
            Value = value.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with an integer value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// </summary>
        public DataCell(int? value)
        {
            DataType = CellDataType.Number;
            Value = value?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a long integer value.
        /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
        /// </summary>
        public DataCell(long value) : this((double)value)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a long integer value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
        /// </summary>
        public DataCell(long? value) : this((double?)value)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a floating point value.
        /// </summary>
        public DataCell(float value)
        {
            DataType = CellDataType.Number;
            Value = value.ToString("G7", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a floating-point value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// </summary>
        public DataCell(float? value)
        {
            DataType = CellDataType.Number;
            Value = value?.ToString("G7", CultureInfo.InvariantCulture) ?? string.Empty;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a double-precision floating-point value.
        /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
        /// </summary>
        public DataCell(double value)
        {
            DataType = CellDataType.Number;
            Value = value.ToString("G15", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a double-precision floating-point value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
        /// </summary>
        public DataCell(double? value)
        {
            DataType = CellDataType.Number;
            Value = value?.ToString("G15", CultureInfo.InvariantCulture) ?? string.Empty;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a decimal floating-point value.
        /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
        /// </summary>
        public DataCell(decimal value) : this(decimal.ToDouble(value))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a decimal floating-point value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
        /// </summary>
        public DataCell(decimal? value) : this(value != null ? decimal.ToDouble(value.Value) : null as double?)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a boolean value.
        /// </summary>
        public DataCell(bool value)
        {
            DataType = CellDataType.Boolean;
            Value = GetStringValue(value);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a boolean value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// </summary>
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
