using System.Buffers;
using System.Buffers.Text;
using System.Drawing;
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
    public static bool TryWrite(uint value, Span<byte> bytes, ref int bytesWritten)
    {
        if (!Utf8Formatter.TryFormat(value, bytes.Slice(bytesWritten), out var length)) return false;
        bytesWritten += length;
        return true;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryWrite(ushort value, Span<byte> bytes, ref int bytesWritten, StandardFormat format = default)
    {
        if (!Utf8Formatter.TryFormat(value, bytes.Slice(bytesWritten), out var length, format)) return false;
        bytesWritten += length;
        return true;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryWrite(double value, Span<byte> bytes, ref int bytesWritten, StandardFormat format = default)
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

#if NETSTANDARD2_0
    public static bool TryWrite(string? value, Span<byte> bytes, ref int bytesWritten)
        => TryWrite(value.AsSpan(), bytes, ref bytesWritten);
#endif

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryWrite(ReadOnlySpan<char> value, Span<byte> bytes, ref int bytesWritten)
    {
        if (!Utf8Helper.TryGetBytes(value, bytes.Slice(bytesWritten), out var length)) return false;
        bytesWritten += length;
        return true;
    }

#if NETSTANDARD2_0
    public static bool TryWriteLongString(string? value, ref int valueIndex, Span<byte> bytes, ref int bytesWritten)
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

    private static bool TryWriteColorChannel(int value, Span<byte> span, ref int written)
        => TryWrite(value, span, ref written, new StandardFormat('X', 2));

    public static bool TryWriteCellReference(int column, uint row, Span<byte> destination, ref int bytesWritten)
    {
        var span = destination.Slice(bytesWritten);

        if (!SpreadsheetUtility.TryGetColumnNameUtf8(column, span, out var written)) return false;
        if (!TryWrite(row, span, ref written)) return false;

        bytesWritten += written;
        return true;
    }
}
