using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal static class Guard
{
    public static TEnum DefinedEnumValue<TEnum>(TEnum value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
        where TEnum : struct, Enum
    {
        if (!EnumPolyfill.IsDefined(value))
            ThrowHelper.EnumValueInvalid(paramName, value);

        return value;
    }

    public static TEnum? DefinedEnumValue<TEnum>(TEnum? value, [CallerArgumentExpression(nameof(value))] string? paramName = null)
        where TEnum : struct, Enum
    {
        if (value is not { } actualValue)
            return null;
        if (!EnumPolyfill.IsDefined(actualValue))
            ThrowHelper.EnumValueInvalid(paramName, actualValue);

        return value;
    }
}
