using SpreadCheetah.Worksheets;
using System.Buffers.Text;

namespace SpreadCheetah.Helpers;

internal static class CellRowHelper
{
    public static int BasicRowStartMaxByteCount { get; } =
        RowStart.Length +
        SpreadsheetConstants.RowIndexMaxDigits +
        RowStartEndTag.Length;

    public static int ConfiguredRowStartMaxByteCount { get; } =
        RowStart.Length +
        SpreadsheetConstants.RowIndexMaxDigits +
        RowHeightStart.Length +
        ValueConstants.DoubleValueMaxCharacters +
        RowHeightEnd.Length;

    public static bool TryWriteRowStart(int rowIndex, SpreadsheetBuffer buffer)
    {
        var bytes = buffer.GetSpan();

        if (RowStart.TryCopyTo(bytes)
            && Utf8Formatter.TryFormat(rowIndex, bytes.Slice(RowStart.Length), out var part2)
            && RowStartEndTag.TryCopyTo(bytes.Slice(RowStart.Length + part2)))
        {
            buffer.Advance(RowStart.Length + RowStartEndTag.Length + part2);
            return true;
        }

        return false;
    }

    public static int GetRowStartBytes(int rowIndex, Span<byte> bytes)
    {
        var bytesWritten = SpanHelper.GetBytes(RowStart, bytes);
        bytesWritten += Utf8Helper.GetBytes(rowIndex, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(RowStartEndTag, bytes.Slice(bytesWritten));
        return bytesWritten;
    }

    public static int GetRowStartBytes(int rowIndex, RowOptions options, Span<byte> bytes)
    {
        if (options.Height is null)
            return GetRowStartBytes(rowIndex, bytes);

        var bytesWritten = SpanHelper.GetBytes(RowStart, bytes);
        bytesWritten += Utf8Helper.GetBytes(rowIndex, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(RowHeightStart, bytes.Slice(bytesWritten));
        bytesWritten += Utf8Helper.GetBytes(options.Height.Value, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(RowHeightEnd, bytes.Slice(bytesWritten));
        return bytesWritten;
    }

    // <row r="
    private static ReadOnlySpan<byte> RowStart => new[]
    {
        (byte)'<', (byte)'r', (byte)'o', (byte)'w', (byte)' ', (byte)'r', (byte)'=', (byte)'"'
    };

    // ">
    private static ReadOnlySpan<byte> RowStartEndTag => new[]
    {
        (byte)'"', (byte)'>'
    };

    // " ht="
    public static ReadOnlySpan<byte> RowHeightStart => new[]
    {
        (byte)'"', (byte)' ', (byte)'h', (byte)'t', (byte)'=', (byte)'"'
    };

    // " customHeight="1">
    public static ReadOnlySpan<byte> RowHeightEnd => new[]
    {
        (byte)'"', (byte)' ', (byte)'c', (byte)'u', (byte)'s', (byte)'t', (byte)'o', (byte)'m', (byte)'H', (byte)'e',
        (byte)'i', (byte)'g', (byte)'h', (byte)'t', (byte)'=', (byte)'"', (byte)'1', (byte)'"', (byte)'>'
    };

    // </row>
    public static ReadOnlySpan<byte> RowEnd => new[]
    {
        (byte)'<', (byte)'/', (byte)'r', (byte)'o', (byte)'w', (byte)'>'
    };
}
