namespace SpreadCheetah.Helpers;

internal static class StringExtensions
{
    public static string? WithEnsuredMaxLength(this string? value, int maxLength) =>
        value is not null && value.Length > maxLength
            ? throw new ArgumentException($"The value can not exceed {maxLength} characters.", nameof(value))
            : value;
}
