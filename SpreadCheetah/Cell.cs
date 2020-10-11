using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;

namespace SpreadCheetah
{
    public readonly struct Cell : IEquatable<Cell>
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

        public override bool Equals(object? obj) => obj is Cell cell && Equals(cell);
        public bool Equals(Cell other) => DataType == other.DataType && Value == other.Value;

        public override int GetHashCode()
        {
            var hashCode = -1382053921;
            hashCode = hashCode * -1521134295 + DataType.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Value);
            return hashCode;
        }

        public static bool operator ==(Cell left, Cell right) => left.Equals(right);
        public static bool operator !=(Cell left, Cell right) => !left.Equals(right);
    }
}
