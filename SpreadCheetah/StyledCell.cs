using SpreadCheetah.Styling;
using System;
using System.Collections.Generic;

namespace SpreadCheetah
{
    public readonly struct StyledCell : IEquatable<StyledCell>
    {
        /// <summary>
        /// Data type and value for the cell.
        /// </summary>
        public DataCell DataCell { get; }

        /// <summary>
        /// Identifier for the style used by this cell. It is optional and can be <c>null</c>.
        /// Style IDs are created with <see cref="Spreadsheet.AddStyle(Style)"/>.
        /// </summary>
        public StyleId? StyleId { get; }

        public StyledCell(string? value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(int value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(int? value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(long value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(long? value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(float value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(float? value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(double value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(double? value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(decimal value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(decimal? value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(bool value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        public StyledCell(bool? value, StyleId? styleId)
        {
            DataCell = new DataCell(value);
            StyleId = styleId;
        }

        /// <inheritdoc/>
        public bool Equals(StyledCell other)
        {
            return DataCell.Equals(other.DataCell) && EqualityComparer<StyleId?>.Default.Equals(StyleId, other.StyleId);
        }

        /// <inheritdoc/>
        public override bool Equals(object? obj) => obj is StyledCell other && Equals(other);

        /// <inheritdoc/>
        public override int GetHashCode() => HashCode.Combine(DataCell, StyleId);

        /// <summary>Determines whether two instances have the same value.</summary>
        public static bool operator ==(StyledCell left, StyledCell right) => left.Equals(right);

        /// <summary>Determines whether two instances have different values.</summary>
        public static bool operator !=(StyledCell left, StyledCell right) => !left.Equals(right);
    }
}
