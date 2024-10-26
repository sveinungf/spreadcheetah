namespace SpreadCheetah.TestHelpers.Helpers;

internal static class EnumHelper
{
    public static bool IsDefined<TEnum>(TEnum value) where TEnum : struct, Enum
    {
#if NET5_0_OR_GREATER
        return Enum.IsDefined(value);
#else
        return Enum.IsDefined(typeof(TEnum), value);
#endif
    }
}
