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
        if (options is null or { DefaultStyleId: null, Height: null })
            return TryWriteRowStart(rowIndex, buffer);

        var written = 0;

        if (!buffer.TryWriteWithoutAdvancing($"{RowStart}{rowIndex}", ref written))
            return false;

        if (options.DefaultStyleId is { } styleId
            && !buffer.TryWriteWithoutAdvancing($"{RowStyleStart}{styleId.Id}{RowStyleEnd}", ref written))
        {
            return false;
        }

        if (options.Height is { } height
            && !buffer.TryWriteWithoutAdvancing($"{RowHeightStart}{height}{RowHeightEnd}", ref written))
        {
            return false;
        }

        if (!buffer.TryWriteWithoutAdvancing(RowStartEndTag, ref written))
            return false;

        buffer.Advance(written);
        return true;
    }

    private static ReadOnlySpan<byte> RowStart => "<row r=\""u8;
    private static ReadOnlySpan<byte> RowStyleStart => "\" s=\""u8;
    private static ReadOnlySpan<byte> RowStyleEnd => "\" customFormat=\"1"u8;
    private static ReadOnlySpan<byte> RowHeightStart => "\" ht=\""u8;
    private static ReadOnlySpan<byte> RowHeightEnd => "\" customHeight=\"1"u8;
    private static ReadOnlySpan<byte> RowStartEndTag => "\">"u8;
}
