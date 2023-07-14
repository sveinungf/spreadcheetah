#if !NET6_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers.Backporting;

internal static class ArgumentNullExceptionBackport
{
    public static void ThrowIfNull([NotNull] object? argument, [CallerArgumentExpression(nameof(argument))] string? paramName = null)
    {
        if (argument is null)
            Throw(paramName);
    }

    [DoesNotReturn]
    private static void Throw(string? paramName) => throw new System.ArgumentNullException(paramName);
}
#endif