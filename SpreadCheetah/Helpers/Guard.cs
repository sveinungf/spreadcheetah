using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal static class Guard
{
    public static double? ColumnWidthInRange(double? width,
        [CallerArgumentExpression(nameof(width))] string? paramName = null)
    {
        if (width is { } value and (< 0 or > 255))
            ThrowHelper.ColumnWidthOutOfRange(paramName, value);

        return width;
    }

    public static TEnum DefinedEnumValue<TEnum>(TEnum value,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
        where TEnum : struct, Enum
    {
        if (!Enum.IsDefined(value))
            ThrowHelper.EnumValueInvalid(paramName, value);

        return value;
    }

    public static TEnum? DefinedEnumValue<TEnum>(TEnum? value,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
        where TEnum : struct, Enum
    {
        if (value is not { } actualValue)
            return null;
        if (!Enum.IsDefined(actualValue))
            ThrowHelper.EnumValueInvalid(paramName, actualValue);

        return value;
    }

    public static double FontSizeInRange(double size,
        [CallerArgumentExpression(nameof(size))] string? paramName = null)
    {
        if (size < 1 || size > 409)
            ThrowHelper.FontSizeOutOfRange(paramName, size);

        return size;
    }

    public static string? FontNameLengthInRange(string? value,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        if (value is null)
            return null;
        if (value.Length < 3)
            ThrowHelper.ValueTooShort(3, paramName);
        if (value.Length > 31)
            ThrowHelper.ValueTooLong(31, paramName);

        return value;
    }

    public static int? FrozenColumnsInRange(int? value,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        const int maxValue = SpreadsheetConstants.MaxNumberOfColumns;
        if (value is < 1 or > maxValue)
            throw new ArgumentOutOfRangeException(paramName, value, $"Number of frozen columns must be between 1 and {maxValue}");

        return value;
    }

    public static int? FrozenRowsInRange(int? value,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        const int maxValue = SpreadsheetConstants.MaxNumberOfRows;
        if (value is < 1 or > maxValue)
            throw new ArgumentOutOfRangeException(paramName, value, $"Number of frozen rows must be between 1 and {maxValue}");

        return value;
    }

    public static string? MaxLength(string? value, int maxLength,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        if (value is not null && value.Length > maxLength)
            ThrowHelper.ValueTooLong(maxLength, paramName);

        return value;
    }

    public static int NotNegative(int value,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        if (value < 0)
            ThrowHelper.ValueIsNegative(paramName, value);

        return value;
    }

    public static double? RowHeightInRange(double? value,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        if (value is <= 0 or > 409)
            throw new ArgumentOutOfRangeException(paramName, value, "Row height must be between 0 and 409.");

        return value;
    }

    public static int SufficientBufferSize(int size,
    [CallerArgumentExpression(nameof(size))] string? paramName = null)
    {
        var minimumSize = SpreadCheetahOptions.MinimumBufferSize;
        if (size < minimumSize)
            throw new ArgumentOutOfRangeException(paramName, size, "Buffer size must be at least " + minimumSize);

        return size;
    }
}
