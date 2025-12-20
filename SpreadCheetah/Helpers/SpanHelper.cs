using System.Buffers;
using System.Buffers.Text;
using System.Drawing;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal static class SpanHelper
{
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

    public static bool TryWrite(Color color, Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!TryWriteColorChannel(color.A, span, ref written)) return false;
        if (!TryWriteColorChannel(color.R, span, ref written)) return false;
        if (!TryWriteColorChannel(color.G, span, ref written)) return false;
        if (!TryWriteColorChannel(color.B, span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private static bool TryWriteColorChannel(int value, Span<byte> span, ref int written)
        => TryWrite(value, span, ref written, new StandardFormat('X', 2));
}
