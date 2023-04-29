using System.Buffers;
using System.Buffers.Text;
using System.Diagnostics;
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

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryWrite(int value, Span<byte> bytes, ref int bytesWritten, StandardFormat format = default)
    {
        if (!Utf8Formatter.TryFormat(value, bytes.Slice(bytesWritten), out var length, format)) return false;
        bytesWritten += length;
        return true;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryWrite(double value, Span<byte> bytes, ref int bytesWritten)
    {
        if (!Utf8Formatter.TryFormat(value, bytes.Slice(bytesWritten), out var length)) return false;
        bytesWritten += length;
        return true;
    }

#if NETSTANDARD2_0
    public static bool TryWrite(string value, Span<byte> bytes, ref int bytesWritten)
        => TryWrite(value.AsSpan(), bytes, ref bytesWritten);
#endif

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryWrite(ReadOnlySpan<char> value, Span<byte> bytes, ref int bytesWritten)
    {
        Debug.Assert(value.Length <= SpreadCheetahOptions.MinimumBufferSize);
        if (!Utf8Helper.TryGetBytes(value, bytes.Slice(bytesWritten), out var length)) return false;
        bytesWritten += length;
        return true;
    }

#if NETSTANDARD2_0
    public static bool TryWriteLongString(string value, ref int valueIndex, Span<byte> bytes, ref int bytesWritten)
        => TryWriteLongString(value.AsSpan(), ref valueIndex, bytes, ref bytesWritten);
#endif

    public static bool TryWriteLongString(ReadOnlySpan<char> value, ref int valueIndex, Span<byte> bytes, ref int bytesWritten)
    {
        if (value.IsEmpty) return true;

        var span = bytes.Slice(bytesWritten);
        var maxCharCount = span.Length / Utf8Helper.MaxBytePerChar;
        var remainingLength = value.Length - valueIndex;
        var lastIteration = remainingLength <= maxCharCount;
        var length = lastIteration ? remainingLength : maxCharCount;
        bytesWritten += Utf8Helper.GetBytes(value.Slice(valueIndex, length), span);
        valueIndex += length;
        return lastIteration;
    }
}
