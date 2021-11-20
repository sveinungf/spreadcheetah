namespace SpreadCheetah.Helpers;

internal static class SpanHelper
{
    public static int GetBytes(ReadOnlySpan<byte> source, Span<byte> destination)
    {
        source.CopyTo(destination);
        return source.Length;
    }
}
