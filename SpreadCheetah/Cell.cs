using SpreadCheetah.Styling;

namespace SpreadCheetah;

/// <summary>
/// Represents the content and an optional style for a worksheet cell.
/// The content can either be a value, or a formula with an optional cached value.
/// Style IDs are created with <see cref="Spreadsheet.AddStyle(Style)"/>.
/// </summary>
public readonly struct Cell : IEquatable<Cell>
{
    internal DataCell DataCell { get; }

    internal Formula? Formula { get; }

    internal StyleId? StyleId { get; }

    private Cell(DataCell dataCell, Formula? formula, StyleId? styleId) => (DataCell, Formula, StyleId) = (dataCell, formula, styleId);

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a text value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// </summary>
    public Cell(string? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with an integer value and an optional style.
    /// </summary>
    public Cell(int value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with an integer value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// </summary>
    public Cell(int? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a long integer value and an optional style.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public Cell(long value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a long integer value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public Cell(long? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a floating point value and an optional style.
    /// </summary>
    public Cell(float value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a floating point value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// </summary>
    public Cell(float? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a double-precision floating-point value and an optional style.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public Cell(double value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a double-precision floating-point value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public Cell(double? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a decimal floating-point value and an optional style.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public Cell(decimal value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a decimal floating-point value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public Cell(decimal? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    public Cell(DateTime value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a boolean value and an optional style.
    /// </summary>
    public Cell(bool value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a boolean value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// </summary>
    public Cell(bool? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula and an optional style.
    /// </summary>
    public Cell(Formula formula, StyleId? styleId = null) : this(formula, null as int?, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a text formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, string? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, int cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, int? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, long cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, long? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, float cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, float? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, double cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, double? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, decimal cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, decimal? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a boolean formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, bool cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a boolean formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, bool? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <inheritdoc/>
    public bool Equals(Cell other)
    {
        return DataCell.Equals(other.DataCell)
            && EqualityComparer<Formula?>.Default.Equals(Formula, other.Formula)
            && EqualityComparer<StyleId?>.Default.Equals(StyleId, other.StyleId);
    }

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is Cell other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(DataCell, Formula, StyleId);

    /// <summary>Determines whether two instances have the same value.</summary>
    public static bool operator ==(in Cell left, in Cell right) => left.Equals(right);

    /// <summary>Determines whether two instances have different values.</summary>
    public static bool operator !=(in Cell left, in Cell right) => !left.Equals(right);
}
