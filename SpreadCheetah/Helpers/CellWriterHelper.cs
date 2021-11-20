namespace SpreadCheetah.Helpers;

internal static class CellWriterHelper
{
    public static readonly int RowStartMaxByteCount = RowStart.Length + SpreadsheetConstants.RowIndexMaxDigits + RowStartEndTag.Length;

    public static ReadOnlySpan<byte> RowStart => new[]
    {
        (byte)'<', (byte)'r', (byte)'o', (byte)'w', (byte)' ', (byte)'r', (byte)'=', (byte)'"'
    };

    public static ReadOnlySpan<byte> RowStartEndTag => new[]
    {
        (byte)'"', (byte)'>'
    };

    public static ReadOnlySpan<byte> RowEnd => new[]
    {
        (byte)'<', (byte)'/', (byte)'r', (byte)'o', (byte)'w', (byte)'>'
    };
}
