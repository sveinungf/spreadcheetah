#if NETSTANDARD2_0
namespace System;

internal static class ReadOnlySpanExtensions
{
    public static int IndexOfAny(this ReadOnlySpan<char> span, string values)
    {
        return span.IndexOfAny(values.AsSpan());
    }
}
#endif