namespace SpreadCheetah.Test.Helpers.Backporting;

internal static class EnumHelper
{
    public static TEnum[] GetValues<TEnum>() where TEnum : struct, Enum
    {
#if NET5_0_OR_GREATER
        return Enum.GetValues<TEnum>();
#else
        return (TEnum[])Enum.GetValues(typeof(TEnum));
#endif
    }
}
