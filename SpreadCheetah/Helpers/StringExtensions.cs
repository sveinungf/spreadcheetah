namespace SpreadCheetah.Helpers;

internal static class StringExtensions
{
    public static bool ContainsChar(this string value, char character)
    {
#if NETSTANDARD2_0
        return value.Contains(character.ToString());
#else
        return value.Contains(character, StringComparison.Ordinal);
#endif
    }
}
