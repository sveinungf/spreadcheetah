using SpreadCheetah.Styling;

namespace SpreadCheetah;

/// <summary>
/// Represents the content and an optional style for a worksheet cell.
/// The content can either be a value, or a formula with an optional cached value.
/// Style IDs are created with <see cref="Spreadsheet.AddStyle(Style)"/>.
/// </summary>
public readonly record struct Cell
{
    internal DataCell DataCell { get; }

    internal Formula? Formula { get; }

    internal StyleId? StyleId { get; }

    private Cell(DataCell dataCell, Formula? formula, StyleId? styleId) => (DataCell, Formula, StyleId) = (dataCell, formula, styleId);

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a text value and an optional style.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public Cell(string? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a text value and an optional style.
    /// </summary>
    public Cell(ReadOnlyMemory<char> value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
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
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
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
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
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
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
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
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
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
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public Cell(decimal? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a <see cref="DateTime"/> value and an optional style.
    /// Will be displayed in the number format from <see cref="Style.Format"/> if set,
    /// otherwise <see cref="SpreadCheetahOptions.DefaultDateTimeFormat"/> will be used instead.
    /// </summary>
    public Cell(DateTime value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a <see cref="DateTime"/> value and an optional style.
    /// Will be displayed in the number format from <see cref="Style.Format"/> if set,
    /// otherwise <see cref="SpreadCheetahOptions.DefaultDateTimeFormat"/> will be used instead.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public Cell(DateTime? value, StyleId? styleId = null) : this(new DataCell(value), null, styleId)
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
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
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
    /// Initializes a new instance of the <see cref="Cell"/> struct with a text formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, ReadOnlyMemory<char> cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
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
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, DateTime cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Cell"/> struct with a number formula, a cached value for the formula, and an optional style.
    /// </summary>
    public Cell(Formula formula, DateTime? cachedValue, StyleId? styleId = null) : this(new DataCell(cachedValue), formula, styleId)
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
}
