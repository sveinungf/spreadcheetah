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
        if (!EnumPolyfill.IsDefined(value))
            ThrowHelper.EnumValueInvalid(paramName, value);

        return value;
    }

    public static TEnum? DefinedEnumValue<TEnum>(TEnum? value,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
        where TEnum : struct, Enum
    {
        if (value is not { } actualValue)
            return null;
        if (!EnumPolyfill.IsDefined(actualValue))
            ThrowHelper.EnumValueInvalid(paramName, actualValue);

        return value;
    }

    public static string? MaxLength(string? value, int maxLength,
        [CallerArgumentExpression(nameof(value))] string? paramName = null)
    {
        if (value is not null && value.Length > maxLength)
            ThrowHelper.ValueTooLong(maxLength, paramName);

        return value;
    }
}
