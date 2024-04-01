using SpreadCheetah.CellValueWriters;

namespace SpreadCheetah;

/// <summary>
/// Represents the value and data type for a worksheet cell.
/// </summary>
public readonly record struct DataCell
{
    internal CellValue NumberValue { get; }
    internal string? StringValue { get; }
    internal CellWriterType Type { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a text value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(string? value)
    {
        StringValue = value;
        Type = value != null ? CellWriterType.String : CellWriterType.Null;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with an integer value.
    /// </summary>
    public DataCell(int value)
    {
        NumberValue = new CellValue(value);
        Type = CellWriterType.Integer;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with an integer value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(int? value)
    {
        NumberValue = value is null ? new CellValue() : new CellValue(value.GetValueOrDefault());
        Type = value is null ? CellWriterType.Null : CellWriterType.Integer;
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
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
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
        NumberValue = new CellValue(value);
        Type = CellWriterType.Float;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a floating-point value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(float? value)
    {
        NumberValue = value is null ? new CellValue() : new CellValue(value.GetValueOrDefault());
        Type = value is null ? CellWriterType.Null : CellWriterType.Float;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a double-precision floating-point value.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public DataCell(double value)
    {
        NumberValue = new CellValue(value);
        Type = CellWriterType.Double;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a double-precision floating-point value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public DataCell(double? value)
    {
        NumberValue = value is null ? new CellValue() : new CellValue(value.GetValueOrDefault());
        Type = value is null ? CellWriterType.Null : CellWriterType.Double;
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
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public DataCell(decimal? value) : this(value != null ? decimal.ToDouble(value.GetValueOrDefault()) : null)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a <see cref="DateTime"/> value.
    /// Will be displayed in the number format from <see cref="SpreadCheetahOptions.DefaultDateTimeFormat"/>.
    /// </summary>
    public DataCell(DateTime value)
    {
        NumberValue = new CellValue(value.ToOADate());
        Type = CellWriterType.DateTime;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a <see cref="DateTime"/> value.
    /// Will be displayed in the number format from <see cref="SpreadCheetahOptions.DefaultDateTimeFormat"/>.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(DateTime? value)
    {
        NumberValue = value is null ? new CellValue() : new CellValue(value.GetValueOrDefault().ToOADate());
        Type = value is null ? CellWriterType.NullDateTime : CellWriterType.DateTime;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a boolean value.
    /// </summary>
    public DataCell(bool value)
    {
        Type = GetBooleanWriter(value);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a boolean value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(bool? value)
    {
        Type = value is null ? CellWriterType.Null : GetBooleanWriter(value.GetValueOrDefault());
    }

    private static CellWriterType GetBooleanWriter(bool value) => value ? CellWriterType.TrueBoolean : CellWriterType.FalseBoolean;
}
