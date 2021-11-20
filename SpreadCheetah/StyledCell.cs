using SpreadCheetah.Styling;

namespace SpreadCheetah;

/// <summary>
/// Represents the value and an optional style for a worksheet cell.
/// Style IDs are created with <see cref="Spreadsheet.AddStyle(Style)"/>.
/// </summary>
public readonly struct StyledCell : IEquatable<StyledCell>
{
    internal DataCell DataCell { get; }

    internal StyleId? StyleId { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a text value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// </summary>
    public StyledCell(string? value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with an integer value and an optional style.
    /// </summary>
    public StyledCell(int value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with an integer value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// </summary>
    public StyledCell(int? value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a long integer value and an optional style.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public StyledCell(long value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a long integer value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public StyledCell(long? value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a floating point value and an optional style.
    /// </summary>
    public StyledCell(float value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a floating point value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// </summary>
    public StyledCell(float? value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a double-precision floating-point value and an optional style.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public StyledCell(double value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a double-precision floating-point value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public StyledCell(double? value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a decimal floating-point value and an optional style.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public StyledCell(decimal value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a decimal floating-point value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public StyledCell(decimal? value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a boolean value and an optional style.
    /// </summary>
    public StyledCell(bool value, StyleId? styleId)
    {
        DataCell = new DataCell(value);
        StyleId = styleId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyledCell"/> struct with a boolean value and an optional style.
    /// If <c>value</c> is <c>null</c>, the cell will be empty.
    /// </summary>
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
    public static bool operator ==(in StyledCell left, in StyledCell right) => left.Equals(right);

    /// <summary>Determines whether two instances have different values.</summary>
    public static bool operator !=(in StyledCell left, in StyledCell right) => !left.Equals(right);
}
