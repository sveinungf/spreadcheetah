namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class ObjectExtensions
{
    public static bool IsEnum<TEnum>(this object value, out TEnum result) where TEnum : struct, Enum
    {
        if (Enum.IsDefined(typeof(TEnum), value))
        {
            result = (TEnum)value;
            return true;
        }

        result = default;
        return false;
    }
}
