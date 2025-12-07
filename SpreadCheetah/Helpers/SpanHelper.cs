using System.Buffers;
using System.Buffers.Text;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Text.Unicode;

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

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryWrite(ReadOnlySpan<char> value, Span<byte> bytes, ref int bytesWritten)
    {
        if (!Utf8Helper.TryGetBytes(value, bytes.Slice(bytesWritten), out var length)) return false;
        bytesWritten += length;
        return true;
    }

    public static bool TryWriteLongString(ReadOnlySpan<char> value, ref int valueIndex, Span<byte> bytes, ref int bytesWritten)
    {
        var source = value.Slice(valueIndex);
        if (source.IsEmpty) return true;

        var destination = bytes.Slice(bytesWritten);
        var status = Utf8.FromUtf16(source, destination, out var charsRead, out var newBytesWritten);
        valueIndex += charsRead;
        bytesWritten += newBytesWritten;
        return status is OperationStatus.Done;
    }

    private static bool TryWriteColorChannel(int value, Span<byte> span, ref int written)
        => TryWrite(value, span, ref written, new StandardFormat('X', 2));

    public static bool TryWriteCellReference(int column, uint row, Span<byte> destination, ref int bytesWritten)
    {
        var span = destination.Slice(bytesWritten);
        if (!SpreadsheetUtility.TryGetColumnNameUtf8(column, span, out var nameLength)) return false;

        span = span.Slice(nameLength);
        if (!Utf8Formatter.TryFormat(row, span, out var rowLength)) return false;

        bytesWritten += nameLength + rowLength;
        return true;
    }
}
