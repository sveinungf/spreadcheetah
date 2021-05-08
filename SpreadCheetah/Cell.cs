using SpreadCheetah.Styling;

namespace SpreadCheetah
{
    /// <summary>
    /// Represents the content and an optional style for a worksheet cell.
    /// The content can either be a value, or a formula with an optional cached value.
    /// Style IDs are created with <see cref="Spreadsheet.AddStyle(Style)"/>.
    /// </summary>
    public readonly struct Cell
    {
        internal DataCell DataCell { get; }

        internal Formula? Formula { get; }

        internal StyleId? StyleId { get; }

        private Cell(DataCell dataCell, Formula? formula, StyleId? styleId) => (DataCell, Formula, StyleId) = (dataCell, formula, styleId);

        public Cell(string? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(int value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(int? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(long value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(long? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(float value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(float? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(double value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(double? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(decimal value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(decimal? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(bool value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(bool? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
        {
        }

        public Cell(Formula formula, StyleId? styleId = null) : this(formula, null as int?, styleId)
        {
        }

        public Cell(Formula formula, string? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, int cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, int? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, long cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, long? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, float cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, float? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, double cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, double? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, decimal cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, decimal? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, bool cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }

        public Cell(Formula formula, bool? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
        {
        }
    }
}
