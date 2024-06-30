using SpreadCheetah.Worksheets;

namespace SpreadCheetah.Helpers;

internal static class CellRowHelper
{
    public static bool TryWriteRowStart(uint rowIndex, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{RowStart}{rowIndex}{RowStartEndTag}");
    }

    public static bool TryWriteRowStart(uint rowIndex, RowOptions options, SpreadsheetBuffer buffer)
    {
        return options.Height is { } height
            ? buffer.TryWrite($"{RowStart}{rowIndex}{RowHeightStart}{height}{RowHeightEnd}")
            : TryWriteRowStart(rowIndex, buffer);
    }

    private static ReadOnlySpan<byte> RowStart => "<row r=\""u8;
    private static ReadOnlySpan<byte> RowStartEndTag => "\">"u8;
    private static ReadOnlySpan<byte> RowHeightStart => "\" ht=\""u8;
    private static ReadOnlySpan<byte> RowHeightEnd => "\" customHeight=\"1\">"u8;
}
