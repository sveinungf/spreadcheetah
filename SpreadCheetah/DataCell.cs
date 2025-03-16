using SpreadCheetah.CellValues;
using SpreadCheetah.CellValueWriters;
using SpreadCheetah.Helpers;
using System.Runtime.InteropServices;

namespace SpreadCheetah;

/// <summary>
/// Represents the value and data type for a worksheet cell.
/// </summary>
[StructLayout(LayoutKind.Auto)]
public readonly record struct DataCell
{
    internal CellValue Value { get; }
    internal CellWriterType Type { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a text value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(string? value)
    {
        Type = (CellWriterType)(value != null ? 1 : 0); // Branchless on .NET 8+
        Value = new CellValue(new StringOrPrimitiveCellValue(value));
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a text value.
    /// </summary>
    public DataCell(ReadOnlyMemory<char> value)
    {
        Type = CellWriterType.ReadOnlyMemoryOfChar;
        Value = new CellValue(value);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with an integer value.
    /// </summary>
    public DataCell(int value)
    {
        Type = CellWriterType.Integer;
        Value = new CellValue(new StringOrPrimitiveCellValue(new PrimitiveCellValue(value)));
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with an integer value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(int? value)
    {
        Type = (CellWriterType)((value.HasValue ? 1 : 0) << 1);
        Value = new CellValue(new StringOrPrimitiveCellValue(new PrimitiveCellValue(value.GetValueOrDefault())));
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
        Type = CellWriterType.Float;
        Value = new CellValue(new StringOrPrimitiveCellValue(new PrimitiveCellValue(value)));
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a floating-point value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(float? value)
    {
        Type = value is null ? CellWriterType.Null : CellWriterType.Float;
        Value = new CellValue(new StringOrPrimitiveCellValue(new PrimitiveCellValue(value.GetValueOrDefault())));
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a double-precision floating-point value.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public DataCell(double value)
    {
        Type = CellWriterType.Double;
        Value = new CellValue(new StringOrPrimitiveCellValue(new PrimitiveCellValue(value)));
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a double-precision floating-point value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// Note that Open XML limits the precision to 15 significant digits for numbers. This could potentially lead to a loss of precision.
    /// </summary>
    public DataCell(double? value)
    {
        Type = (CellWriterType)((value.HasValue ? 1 : 0) << 2);
        Value = new CellValue(new StringOrPrimitiveCellValue(new PrimitiveCellValue(value.GetValueOrDefault())));
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
        if (value.Ticks < OADate.MinTicks)
            ThrowHelper.InvalidOADate();

        Type = CellWriterType.DateTime;
        Value = new CellValue(new StringOrPrimitiveCellValue(new PrimitiveCellValue(value.Ticks)));
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a <see cref="DateTime"/> value.
    /// Will be displayed in the number format from <see cref="SpreadCheetahOptions.DefaultDateTimeFormat"/>.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(DateTime? value)
    {
        var ticks = value.GetValueOrDefault().Ticks;
        if (value.HasValue && ticks < OADate.MinTicks)
            ThrowHelper.InvalidOADate();

        Type = (CellWriterType)((value.HasValue ? 1 : 0) + 5);
        Value = new CellValue(new StringOrPrimitiveCellValue(new PrimitiveCellValue(ticks)));
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a boolean value.
    /// </summary>
    public DataCell(bool value)
    {
        Type = GetBooleanWriter(value);
#if NET5_0_OR_GREATER
        System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
#endif
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DataCell"/> struct with a boolean value.
    /// If <c>value</c> is <see langword="null"/>, the cell will be empty.
    /// </summary>
    public DataCell(bool? value)
    {
        Type = value is null ? CellWriterType.Null : GetBooleanWriter(value.GetValueOrDefault());
#if NET5_0_OR_GREATER
        System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
#endif
    }

    private static CellWriterType GetBooleanWriter(bool value) => (CellWriterType)((value ? 1 : 0) + 7);
}
