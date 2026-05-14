using SpreadCheetah.MetadataXml.Attributes;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.Helpers;

internal static class CellRowHelper
{
    public static bool TryWriteRowStart(uint rowIndex, SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite($"{"<row r=\""u8}{rowIndex}{"\">"u8}");
    }

    public static bool TryWriteRowStart(uint rowIndex,
        RowOptions options, StyleId? rowStyleId, SpreadsheetBuffer buffer)
    {
        if (options is null or { DefaultStyle: null, Height: null, OutlineLevel: null, Collapsed: null, Hidden: null })
            return TryWriteRowStart(rowIndex, buffer);

        var sAttribute = new IntAttribute("s"u8, rowStyleId?.Id);
        var customFormatAttribute = new BooleanAttribute("customFormat"u8, rowStyleId is not null ? true : null);
        var htAttribute = new DoubleAttribute("ht"u8, options.Height);
        var customHeightAttribute = new BooleanAttribute("customHeight"u8, options.Height is not null ? true : null);
        var outlineLevelAttribute = new IntAttribute("outlineLevel"u8, options.OutlineLevel);
        var hiddenAttribute = new BooleanAttribute("hidden"u8, options.Hidden);
        var collapsedAttribute = new BooleanAttribute("collapsed"u8, options.Collapsed);

        return buffer.TryWrite(
            $"{"<row r=\""u8}" +
            $"{rowIndex}" +
            $"{"\""u8}" +
            $"{sAttribute}" +
            $"{customFormatAttribute}" +
            $"{htAttribute}" +
            $"{customHeightAttribute}" +
            $"{outlineLevelAttribute}" +
            $"{hiddenAttribute}" +
            $"{collapsedAttribute}" +
            $"{">"u8}");
    }
}
