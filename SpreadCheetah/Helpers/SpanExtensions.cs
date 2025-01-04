namespace SpreadCheetah.Helpers;

internal static class SpanExtensions
{
    public static bool TrySlice<T>(this ReadOnlySpan<T> span, int start, out ReadOnlySpan<T> slice)
    {
        if (start >= 0 && start < span.Length)
        {
            slice = span.Slice(start);
            return true;
        }
        slice = [];
        return false;
    }
}
