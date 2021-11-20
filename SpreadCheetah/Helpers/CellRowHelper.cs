namespace SpreadCheetah.Helpers;

internal static class CellRowHelper
{
    public static readonly int BasicRowStartMaxByteCount =
        RowStart.Length +
        SpreadsheetConstants.RowIndexMaxDigits +
        RowStartEndTag.Length;

    public static int GetRowStartBytes(int rowIndex, Span<byte> bytes)
    {
        var bytesWritten = SpanHelper.GetBytes(RowStart, bytes);
        bytesWritten += Utf8Helper.GetBytes(rowIndex, bytes.Slice(bytesWritten));
        bytesWritten += SpanHelper.GetBytes(RowStartEndTag, bytes.Slice(bytesWritten));
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

    // </row>
    public static ReadOnlySpan<byte> RowEnd => new[]
    {
        (byte)'<', (byte)'/', (byte)'r', (byte)'o', (byte)'w', (byte)'>'
    };
}
