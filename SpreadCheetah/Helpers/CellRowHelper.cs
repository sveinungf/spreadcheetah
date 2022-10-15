using SpreadCheetah.Worksheets;
using System.Buffers.Text;

namespace SpreadCheetah.Helpers;

internal static class CellRowHelper
{
    public static int BasicRowStartMaxByteCount { get; } =
        RowStart.Length +
        SpreadsheetConstants.RowIndexMaxDigits +
        RowStartEndTag.Length;

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

    public static bool TryWriteRowStart(int rowIndex, RowOptions options, SpreadsheetBuffer buffer)
    {
        if (options.Height is null)
            return TryWriteRowStart(rowIndex, buffer);

        var bytes = buffer.GetSpan();
        var part1 = RowStart.Length;
        var part3 = RowHeightStart.Length;
        var part5 = RowHeightEnd.Length;

        if (RowStart.TryCopyTo(bytes)
            && Utf8Formatter.TryFormat(rowIndex, bytes.Slice(part1), out var part2)
            && RowHeightStart.TryCopyTo(bytes.Slice(part1 + part2))
            && Utf8Formatter.TryFormat(options.Height.Value, bytes.Slice(part1 + part2 + part3), out var part4)
            && RowHeightEnd.TryCopyTo(bytes.Slice(part1 + part2 + part3 + part4)))
        {
            buffer.Advance(part1 + part2 + part3 + part4 + part5);
            return true;
        }

        return false;
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
