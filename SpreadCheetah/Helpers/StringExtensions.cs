using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal static class StringExtensions
{
    // Invalid worksheet name characters in Excel
    private const string InvalidSheetNameCharacters = @"/\*?[]";

#if NET8_0_OR_GREATER
    private static readonly System.Buffers.SearchValues<char> InvalidSheetNameCharactersSearchValues = System.Buffers.SearchValues.Create(InvalidSheetNameCharacters);
#endif

    public static string? WithEnsuredMaxLength(this string? value, int maxLength) =>
        value is not null && value.Length > maxLength
            ? throw new ArgumentException(StringHelper.Invariant($"The value can not exceed {maxLength} characters."), nameof(value))
            : value;

    public static void EnsureValidWorksheetName(this string name, [CallerArgumentExpression(nameof(name))] string? paramName = null)
    {
        if (string.IsNullOrWhiteSpace(name))
            ThrowHelper.NameEmptyOrWhiteSpace(paramName);

        if (name.Length > 31)
            ThrowHelper.WorksheetNameTooLong(paramName);

        if (name.StartsWith('\'') || name.EndsWith('\''))
            ThrowHelper.WorksheetNameStartsOrEndsWithSingleQuote(paramName);

#if NET8_0_OR_GREATER
        if (name.AsSpan().ContainsAny(InvalidSheetNameCharactersSearchValues))
#else
        if (name.AsSpan().IndexOfAny(InvalidSheetNameCharacters) != -1)
#endif
        {
            ThrowHelper.WorksheetNameInvalidCharacters(paramName, InvalidSheetNameCharacters);
        }
    }
}
