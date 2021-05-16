using SpreadCheetah.CellValueWriters;
using System;
using System.Net;

namespace SpreadCheetah
{
    /// <summary>
    /// Represents the value and data type for a worksheet cell.
    /// </summary>
    public readonly struct DataCell : IEquatable<DataCell>
    {
        internal CellValue NumberValue { get; }

        internal string? StringValue { get; }

        internal CellValueWriter Writer { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a text value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// </summary>
        public DataCell(string? value) : this()
        {
            StringValue = value != null ? WebUtility.HtmlEncode(value) : string.Empty;
            Writer = value != null ? CellValueWriter.String : CellValueWriter.Null;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with an integer value.
        /// </summary>
        public DataCell(int value) : this()
        {
            NumberValue = new CellValue(value);
            Writer = CellValueWriter.Integer;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with an integer value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// </summary>
        public DataCell(int? value) : this()
        {
            NumberValue = value is null ? new CellValue() : new CellValue(value.Value);
            Writer = value is null ? CellValueWriter.Null : CellValueWriter.Integer;
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
        public DataCell(float value) : this()
        {
            NumberValue = new CellValue(value);
            Writer = CellValueWriter.Float;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a floating-point value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// </summary>
        public DataCell(float? value) : this()
        {
            NumberValue = value is null ? new CellValue() : new CellValue(value.Value);
            Writer = value is null ? CellValueWriter.Null : CellValueWriter.Float;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a double-precision floating-point value.
        /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
        /// </summary>
        public DataCell(double value) : this()
        {
            NumberValue = new CellValue(value);
            Writer = CellValueWriter.Double;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a double-precision floating-point value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
        /// </summary>
        public DataCell(double? value) : this()
        {
            NumberValue = value is null ? new CellValue() : new CellValue(value.Value);
            Writer = value is null ? CellValueWriter.Null : CellValueWriter.Double;
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
        public DataCell(decimal? value) : this(value != null ? decimal.ToDouble(value.Value) : null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a boolean value.
        /// </summary>
        public DataCell(bool value) : this()
        {
            Writer = GetBooleanWriter(value);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> struct with a boolean value.
        /// If <c>value</c> is <c>null</c>, the cell will be empty.
        /// </summary>
        public DataCell(bool? value) : this()
        {
            Writer = value is null ? CellValueWriter.Null : GetBooleanWriter(value.Value);
        }

        private static CellValueWriter GetBooleanWriter(bool value) => value ? CellValueWriter.TrueBoolean : CellValueWriter.FalseBoolean;

        /// <inheritdoc/>
        public bool Equals(DataCell other) => Writer == other.Writer
            && string.Equals(StringValue, other.StringValue, StringComparison.Ordinal)
            && Writer.Equals(NumberValue, other.NumberValue);

        /// <inheritdoc/>
        public override bool Equals(object? obj) => obj is DataCell cell && Equals(cell);

        /// <inheritdoc/>
        public override int GetHashCode() => HashCode.Combine(StringValue, Writer, Writer.GetHashCodeFor(NumberValue));

        /// <summary>Determines whether two instances have the same value.</summary>
        public static bool operator ==(DataCell left, DataCell right) => left.Equals(right);

        /// <summary>Determines whether two instances have different values.</summary>
        public static bool operator !=(DataCell left, DataCell right) => !left.Equals(right);
    }
}
