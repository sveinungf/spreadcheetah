#if NETSTANDARD2_0
namespace System;

internal static class StringExtensions
{
#pragma warning disable IDE0060, RCS1163 // Unused parameter.
    public static bool Contains(this string value, char character, StringComparison comparisonType)
#pragma warning restore IDE0060, RCS1163 // Unused parameter.
    {
        return value.IndexOf(character) >= 0;
    }

    public static bool StartsWith(this string value, char character)
    {
        return value.Length > 0 && value[0] == character;
    }

    public static bool EndsWith(this string value, char character)
    {
        return value.Length > 0 && value[value.Length - 1] == character;
    }
}
#endif