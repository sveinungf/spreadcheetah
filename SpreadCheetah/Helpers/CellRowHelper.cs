using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.Helpers;

internal static class CellRowHelper
{
    public static bool TryWriteRowStart(uint rowIndex, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{RowStart}{rowIndex}{RowStartEndTag}");
    }

    public static bool TryWriteRowStart(uint rowIndex,
        RowOptions options, StyleId? rowStyleId, SpreadsheetBuffer buffer)
    {
        if (options is null or { DefaultStyle: null, Height: null })
            return TryWriteRowStart(rowIndex, buffer);

        var indexBefore = buffer.Index;

        if (!buffer.TryWrite($"{RowStart}{rowIndex}"))
            return Fail();

        if (rowStyleId is { } styleId
            && !buffer.TryWrite($"{RowStyleStart}{styleId.Id}{RowStyleEnd}"))
        {
            return Fail();
        }

        if (options.Height is { } height
            && !buffer.TryWrite($"{RowHeightStart}{height}{RowHeightEnd}"))
        {
            return Fail();
        }

        return buffer.TryWrite(RowStartEndTag) || Fail();

        bool Fail()
        {
            buffer.Advance(indexBefore - buffer.Index);
            return false;
        }
    }

    private static ReadOnlySpan<byte> RowStart => "<row r=\""u8;
    private static ReadOnlySpan<byte> RowStyleStart => "\" s=\""u8;
    private static ReadOnlySpan<byte> RowStyleEnd => "\" customFormat=\"1"u8;
    private static ReadOnlySpan<byte> RowHeightStart => "\" ht=\""u8;
    private static ReadOnlySpan<byte> RowHeightEnd => "\" customHeight=\"1"u8;
    private static ReadOnlySpan<byte> RowStartEndTag => "\">"u8;
}
