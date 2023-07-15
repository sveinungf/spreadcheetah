using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal static class StringHelper
{
#if NET6_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string Invariant(ref DefaultInterpolatedStringHandler handler) => string.Create(System.Globalization.CultureInfo.InvariantCulture, ref handler);
#else
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string Invariant(FormattableString formattable) => FormattableString.Invariant(formattable);
#endif
}
