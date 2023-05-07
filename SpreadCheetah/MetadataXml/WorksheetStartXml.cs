using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.MetadataXml;

internal struct WorksheetStartXml
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""u8;

    private static ReadOnlySpan<byte> SheetDataBegin => "<sheetData>"u8;

    private readonly WorksheetOptions? _options;
    private Element _next;

    public WorksheetStartXml(WorksheetOptions? options)
    {
        _options = options;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.SheetViews && !Advance(TryWriteSheetViews(bytes, ref bytesWritten))) return false;

        // TODO
        if (_options is null)
        {
            if (_next == Element.SheetDataBegin && !Advance(SheetDataBegin.TryCopyTo(bytes, ref bytesWritten))) return false;
        }

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteSheetViews(Span<byte> bytes, ref int bytesWritten)
    {
        var options = _options;
        if (options?.FrozenColumns is null && options?.FrozenRows is null)
            return true;

        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"<sheetViews><sheetView workbookViewId=\"0\"><pane "u8.TryCopyTo(span, ref written)) return false;

        if (options.FrozenColumns is not null)
        {
            if (!"xSplit=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(options.FrozenColumns.Value, span, ref written)) return false;
            if (!"\" "u8.TryCopyTo(span, ref written)) return false;
        }

        if (options.FrozenRows is not null)
        {
            if (!"ySplit=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(options.FrozenRows.Value, span, ref written)) return false;
            if (!"\" "u8.TryCopyTo(span, ref written)) return false;
        }

        if (!"topLeftCell=\""u8.TryCopyTo(span, ref written)) return false;

        // TODO: Make overload of GetColumnName that writes into span
        var columnName = SpreadsheetUtility.GetColumnName((options.FrozenColumns ?? 0) + 1);
        if (!SpanHelper.TryWrite(columnName, span, ref written)) return false;
        if (!SpanHelper.TryWrite((options.FrozenRows ?? 0) + 1, span, ref written)) return false;
        if (!"\" activePane=\""u8.TryCopyTo(span, ref written)) return false;

        var activePane = options switch
        {
            { FrozenColumns: not null, FrozenRows: not null } => "bottomRight"u8,
            { FrozenColumns: not null } => "topRight"u8,
            _ => "bottomLeft"u8
        };

        if (!activePane.TryCopyTo(span, ref written)) return false;
        if (!"\" state=\"frozen\"/>"u8.TryCopyTo(span, ref written)) return false;

        var bottomRight = """<selection pane="bottomRight"/>"""u8;
        if (options is { FrozenColumns: not null, FrozenRows: not null } && !bottomRight.TryCopyTo(span, ref written))
            return false;

        var bottomLeft = """<selection pane="bottomLeft"/>"""u8;
        if (options.FrozenRows is not null && !bottomLeft.TryCopyTo(span, ref written))
            return false;

        var topRight = """<selection pane="topRight"/>"""u8;
        if (options.FrozenColumns is not null && !topRight.TryCopyTo(span, ref written))
            return false;

        if (!"</sheetView></sheetViews>"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private enum Element
    {
        Header,
        SheetViews,
        SheetDataBegin,
        Done
    }
}
