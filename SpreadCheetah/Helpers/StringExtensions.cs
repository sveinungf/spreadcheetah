namespace SpreadCheetah.Helpers;

internal static class StringExtensions
{
    public static bool ContainsChar(this string value, char character)
    {
#if NETSTANDARD2_0
        return value.IndexOf(character) >= 0;
#else
        return value.Contains(character, StringComparison.Ordinal);
#endif
    }

    public static bool StartsWithChar(this string value, char character)
    {
#if NETSTANDARD2_0
        return value.Length > 0 && value[0] == character;
#else
        return value.StartsWith(character);
#endif
    }

    public static bool EndsWithChar(this string value, char character)
    {
#if NETSTANDARD2_0
        return value.Length > 0 && value[value.Length - 1] == character;
#else
        return value.EndsWith(character);
#endif
    }

    public static string? WithEnsuredMaxLength(this string? value, int maxLength) =>
        value is not null && value.Length > maxLength
            ? throw new ArgumentException($"The value can not exceed {maxLength} characters.", nameof(value))
            : value;
}
