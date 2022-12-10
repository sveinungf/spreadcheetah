using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal static class StringExtensions
{
    // Invalid worksheet name characters in Excel
    private const string InvalidSheetNameCharacters = @"/\*?[]";

    public static string? WithEnsuredMaxLength(this string? value, int maxLength) =>
        value is not null && value.Length > maxLength
            ? throw new ArgumentException($"The value can not exceed {maxLength} characters.", nameof(value))
            : value;

    public static void EnsureValidWorksheetName(this string name, [CallerArgumentExpression("name")] string? paramName = null)
    {
        if (string.IsNullOrWhiteSpace(name))
            ThrowHelper.WorksheetNameEmptyOrWhiteSpace(paramName);

        if (name.Length > 31)
            ThrowHelper.WorksheetNameTooLong(paramName);

        if (name.StartsWith('\'') || name.EndsWith('\''))
            ThrowHelper.WorksheetNameStartsOrEndsWithSingleQuote(paramName);

        if (name.AsSpan().IndexOfAny(InvalidSheetNameCharacters) != -1)
            ThrowHelper.WorksheetNameInvalidCharacters(paramName, InvalidSheetNameCharacters);
    }
}
