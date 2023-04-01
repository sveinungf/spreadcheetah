using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal static class SpanHelper
{
    public static int GetBytes(ReadOnlySpan<byte> source, Span<byte> destination)
    {
        source.CopyTo(destination);
        return source.Length;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryCopyTo(this ReadOnlySpan<byte> source, Span<byte> bytes, ref int bytesWritten)
    {
        if (!source.TryCopyTo(bytes.Slice(bytesWritten))) return false;
        bytesWritten += source.Length;
        return true;
    }
}
